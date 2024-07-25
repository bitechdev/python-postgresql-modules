from ..exception import MergeToolWarning, MergeToolError
from ...util.object import getProp
from ...util.json import jsoncomplex
from ...util.string import (
    findWordArticle,
    html_replace,
    n2w,
    num2words,
    ordinalstr,
    istrreplace,
)
from ...util.number import ndec, nint, isFloat, tryInt, tryFloat, getNumber
from ...util.array import avg_any, max_any, min_any, sum_any
from ...pgproc.debug import sqlmsg, get_err_msg
from ..common.base import MergeTool

import xml.etree.ElementTree as ET
import lxml.etree as LET
import base64
import ast
import re
import locale


import io, sys, os, traceback
import json
import copy
import datetime
import math


class MergeToolHTML(MergeTool):
    """HTML Merge Tool"""

    def __init__(s):
        super().__init__()

        s.srcdoc = None
        s.destdoc = None
        s.srcfilename = None
        s.destfilename = None
        s.visit_tables_cnt = 0
        s.tbl_basic_replace_cnt = 0
        s.tbl_map = {}
        s.tbl_ordinaldata = {}
        s.tbl_rmrow = []
        s.tbl_rmcol = []
        s.tbl_complex_rowdat = {}
        s.tbl_mode = {}
        s.HTML = True
        s.htmltree = None
        s.parser = None

    #

    def execute(s):
        if s.mode.find("peek") >= 0:
            tagdata = s.peek_html()
            p_result = (
                json.dumps(tagdata, ensure_ascii=False).encode("utf8").decode("utf8")
            )
        #
        elif s.mode.find("convert") >= 0:
            s.convert_html()
            if len(s.faults) > 0:
                p_retval = 2
                p_errmsg = "The following convert fields are not valid: \r\n"
                for f in s.faults:
                    p_errmsg = p_errmsg + "\r\n{}".format(f.get("tag", ""))
                #
            #

        elif s.mode.find("merge") >= 0:
            s.merge_html()
        elif s.mode.find("mocktest") >= 0:
            s.generate_mockup()
        elif s.mode.find("mark") >= 0:
            s.mark_html()

        else:
            raise MergeToolError(
                "Mode {} not supported.".format(s.mode), "MergeToolHTML.init"
            )
        #

        return s.results

    def load_html(s):
        s.parser = LET.HTMLParser(
            recover=True,
            strip_cdata=False,
            no_network=False,
            huge_tree=True,
            compact=False,
        )
        if s.srcfilename is None or len(s.srcfilename) < 3:
            stream = io.BytesIO(s.template_blob)
            if s.template_blob is None:
                raise Exception("Template blob is none")
            #
            s.htmltree = LET.parse(stream, s.parser)

        else:
            s.htmltree = LET.parse(s.srcfilename, s.parser)
        #

    #

    def save_html(s):
        # Save tree
        content = LET.tostring(s.htmltree, pretty_print=True, method="html")

        if s.toplaintext:
            import html2text

            h = html2text.HTML2Text()
            h.ignore_links = True
            txt = h.handle(content.decode("utf8"))
            content = txt.encode("utf8")
        #

        if s.destfilename is None or len(s.destfilename) < 3:
            s.results = content
        else:
            with open(s.destfilename, "wb") as f:
                f.write(content)
            #
            if sys.platform.lower() == "linux":
                os.chmod(s.destfilename, 0o777)
            #
        #
        return s.results

    #

    def peek_html(s):
        """Scans through a string for merge tags and returns a list of tags found within merge tags."""
        tags = {}
        tags["doctags"] = []
        tags["pictures"] = []
        tags["headertags"] = []
        tags["footertags"] = []

        s.load_html()
        if s.htmltree is None:
            raise Exception("No parser")
        #

        tags["tabletags"] = s.visit_tables(s.htmltree)

        for tag in s.htmltree.iter():
            if tag.text is not None:
                tags["doctags"].extend(s.peek_str(tag.text))
            #
            if tag.tail is not None:
                tags["doctags"].extend(s.peek_str(tag.tail))
            #
            if tag.attrib:
                for attrib in tag.attrib:
                    val = tag.get(attrib)
                    if val.find(s.TAG_S) >= 0:
                        tags["doctags"].extend(s.peek_str(val))
                    #
                #
            #

        #

        return tags

    #

    def lxmldepth(s, node, tagtype=None):
        d = 0
        while node is not None:
            if tagtype is None:
                d += 1
            elif node.tag == tagtype:
                d += 1
            #
            node = node.getparent()
        return d

    #

    def visit_tables(s, p_tree, p_parent=0):

        table_tags = {}
        if p_parent == 0:
            s.visit_tables_cnt = 0
        # plpy.notice("Visit table, {}".format(p_tree))

        lastdepth = 0

        def loopanon(src):
            nonlocal table_tags
            # plpy.notice("loopanon src, {} = {}".format(src.tag, src.text))
            for anon in src.getchildren():
                # plpy.notice("loopanon, {} = {}".format(anon.tag, anon.text))
                if anon.tag == "table":
                    subdepth = s.lxmldepth(anon, "table")
                    if subdepth >= depth:
                        tags = s.visit_tables(src, lastid)
                        for t in tags:
                            table_tags[t] = tags[t]
                        #
                    #
                    continue
                #

                loopanon(anon)
            #

            lst = s.peek_str(src.text)
            if lst is not None and len(lst) > 0:
                if table_tags.get("{},{}".format(lastid, p_parent)) is None:
                    table_tags["{},{}".format(lastid, p_parent)] = []
                #
                table_tags["{},{}".format(lastid, p_parent)].extend(lst)
            #

        #

        for tag in p_tree.iter():

            if tag.tag == "table":
                s.visit_tables_cnt += 1
                depth = s.lxmldepth(tag, "table")
                if lastdepth == 0:
                    lastdepth = depth
                elif (
                    depth > lastdepth
                ):  # this is to skip the children inside children and instead call them below
                    continue
                #
                # plpy.notice("Table Depth, {}".format(depth))
                table_tags["{},{}".format(s.visit_tables_cnt, p_parent)] = []
                lastid = s.visit_tables_cnt

                for row in tag.iterfind("tr"):
                    for cel in row.findall("td"):
                        loopanon(cel)
                    #
                #
                for body in tag.iterfind("tbody"):
                    for row in body.iterfind("tr"):
                        for cel in row.findall("td"):
                            loopanon(cel)
                        #
                    #
                #
            #
        #

        # sqlmsg("Table Tags={}".format( table_tags))
        return table_tags

    #

    def merge_html(s):
        # We are going to use https://lxml.de/ which can parse html and already used for docx
        s.load_html()

        # dat = s.str_value_replace(content, None, s.inputdata.get('fields'))
        # content = dat["str"]
        # sqlmsg("s.inputdata={}".format(s.inputdata))

        # for tag in htmltree.iter():
        #  plpy.notice("Test Tag={} Path={} Text={} Tail= {}".format(tag.tag, htmltree.getpath(tag),tag.text, tag.tail))
        #

        tables = []
        for tag in s.looptables(s.htmltree.getroot()):
            tables.append(tag)
            # plpy.notice("Tab Loop Tag: ", tag)
        #

        if len(tables) > 0:
            s.tbl_basic_replace(tables)
            s.tbl_complex_init()
            s.tbl_post_loop(tables)
        #

        # Merge leftover tags
        for tag in s.htmltree.iter():
            # plpy.notice("Tag={} Path={} Text={}".format(tag.tag, htmltree.getpath(tag),tag.text))
            # sqlmsg("Tag={} Path={} Text={}".format(tag.tag, s.htmltree.getpath(tag),tag.text))
            s.htmlBasicTagReplace(tag)
        #

        eval_tag_start = 0
        eval_tag_end = 0

        # Save tree
        s.save_html()
        # if sys.platform.lower() == "linux":
        #  os.chmod(s.srcfilename, 0o777)
        #

    #

    def htmlBasicTagReplace(s, tag):

        dat = s.tag_value_replace(tag, None, s.inputdata.get("fields"))

        if tag.text is not None and tag.text != "":
            tag.text = s.eval_replace(tag.text, "html")
            tag.text = s.cleanuptags(tag.text)
        #

        if tag.tail is not None and tag.tail != "":
            tag.tail = s.eval_replace(tag.tail, "html")
            tag.tail = s.cleanuptags(tag.tail)
        #

        for attrib in tag.attrib:
            val = tag.get(attrib)
            if val.find(s.TAG_S) >= 0:
                dat = s.tag_value_replace(val, None, s.inputdata.get("fields"))
                # dat = s.str_value_replace(str(val), None, s.inputdata.get('fields'))
                # sqlmsg("Get attrib {}={} to {} Fields: {}".format(attrib,val, dat , s.inputdata.get('fields')), p_type="local error")
                val = dat["str"]
                val = s.cleanuptags(val)
                tag.set(attrib, val)
            #
        #

    #

    def cleanuptags(s, p_text):
        m_text = p_text
        for i in range(1, 50):
            start = m_text.find(s.TAG_S)
            if start < 0:
                break
            if start >= 0:
                end = m_text.find(s.TAG_E)
                if end < 0:
                    break
                m_tag = m_text[start : end + len(s.TAG_E)]
                if m_tag.lower().find(s.STR_TAG_ESIGN) >= 0:
                    break
                #
                m_text = m_text[:start] + m_text[end + len(s.TAG_E) :]
                plpy.notice("Cleaned Empty Leftover tag: {}".format(m_tag))

                # sqlmsg("Cleaned Empty Leftover tag: {}".format(m_tag))
            #
        #
        m_text = s.cleanImageTags(m_text, 0)
        return m_text

    #

    def cleanImageTags(s, pSrc, pStart=0):
        i_sts = pSrc.find("[^", pStart)
        i_ste = pSrc.find("^]", i_sts)
        if i_sts > 0 and i_ste > i_sts:
            i_ste = i_ste + 2
            pSrc = pSrc[i_sts:] + pSrc[i_ste:]
            sqlmsg(
                "Cleaned {} in {} ".format(pSrc[i_sts:i_ste], pSrc),
                p_title="pl_mailmerge",
                p_type="local notice",
            )
        #
        return pSrc

    #

    def tag_value_replaceAll(
        s, p_xmlnode, p_idx=None, p_fields=None, p_type="", p_tblid="", p_colidx=0
    ):
        opdata = {}
        childcnt = 0
        for subtag in p_xmlnode.iter():
            # plpy.notice("Tag: {} Path: {} Text:{}".format(tag.tag, htmltree.getpath(tag),tag.text))
            if (subtag.text and subtag.text != "") or (
                subtag.tail is not None and subtag.tail != ""
            ):
                # dat = s.str_value_replace(subtag.text, None, s.inputdata.get('fields'),p_htmlnode=subtag)
                childcnt = childcnt + 1
                opdata = s.tag_value_replace(
                    subtag, p_idx, p_fields, p_type, p_tblid, p_colidx
                )
                # subtag.text = dat["str"]
                # subtag.text = s.eval_replace(dat["str"], "html")
            #
        #

        for attrib in p_xmlnode.attrib:
            val = p_xmlnode.get(attrib)
            dat = s.str_value_replace(
                val, None, s.inputdata.get("fields"), p_htmlnode=attrib
            )
            val = dat["str"]
            p_xmlnode.set(attrib, val)
        #

        # if childcnt == 0:
        if (p_xmlnode.text and p_xmlnode.text != "") or (
            p_xmlnode.tail is not None and p_xmlnode.tail != ""
        ):
            opdata = s.tag_value_replace(
                p_xmlnode, p_idx, p_fields, p_type, p_tblid, p_colidx
            )
        #
        return opdata

    #

    def tag_value_replace_node(
        s,
        p_xmlnode,
        p_xmltext,
        p_idx=None,
        p_fields=None,
        p_type="",
        p_tblid="",
        p_colidx=0,
    ):
        """Scans through html objects and replace from the given values."""
        opdata = {
            "error": None,
            "rmblnk": 0,
            "rmdr": 0,
            "mergedtags": [],
            "value": None,
            "rmblnkcol": 0,
            "limitrow": 0,
            "xmltext": p_xmltext,
        }
        isrun = False

        try:

            if p_idx is None:
                p_idx = 0

            if p_xmltext is None:
                opdata["error"] = "No xml object. "
                return opdata
            #

            if (
                not p_xmltext.find(s.TAG_S) >= 0
            ):  # abort this line/object, we have not tags.
                opdata["error"] = "No tags found."
                return opdata
            #

            fields = p_fields
            if fields is None:
                fields = s.inputdata.get("fields")
            #

            if p_xmltext.find(s.TAG_S + s.STR_TAG_RMBD + s.TAG_E) >= 0:
                opdata["rmblnk"] = 1
            #

            if p_xmltext.find(s.TAG_S + s.STR_TAG_RMBC + s.TAG_E) >= 0:
                opdata["rmblnkcol"] = 1
            #

            if p_xmltext.find(s.TAG_S + s.STR_TAG_RMDR + s.TAG_E) >= 0:
                opdata["rmdr"] = 1
            #

            if (
                p_xmltext is not None
                and p_xmltext.find(s.TAG_S + s.STR_TAG_LIMITROW) >= 0
            ):
                istart = p_xmltext.find(s.TAG_S + s.STR_TAG_LIMITROW)
                iend = p_xmltext.find(s.TAG_E, istart)
                parts = (p_xmltext[istart + len(s.TAG_S) : iend]).split(":")
                if len(parts) > 1:
                    opdata["limitrow"] = tryInt(parts[1])
                else:
                    opdata["limitrow"] = 1
                #

                p_xmltext = p_xmltext[:istart] + p_xmltext[iend + len(s.TAG_E) :]

            #

            if fields is None or type(fields) is not dict:
                opdata["error"] = "No fields given."
                return opdata
            #

            lasttag = None
            for tag in fields:
                found = False
                lasttag = tag
                val = None
                tagtype = None
                tagvalue = None
                if tag is None:
                    continue

                try:
                    if type(fields.get(tag)) is not dict:
                        raise MergeToolError(
                            "Expecting tag field to have a type and value property.",
                            "tag_value_replace",
                        )
                    #
                    # if nodevalue is None and p_xmlnode.tail is None : break

                    tagtype = int(fields.get(tag).get("type", 0))
                    tagvalue = fields.get(tag).get("value", "")

                    if (
                        type(tagvalue) is list
                    ):  # we can look at the type also, but sigle values wont have a list
                        try:
                            val = tagvalue[p_idx]
                            found = True
                        except Exception as e:
                            s.add_fault("No data for field {}".format(tagtype), str(e))
                            # plpy.warning("Field replace Exception: {} Detail: i:{} tag:{},datalist:{}".format(e,p_idx, tagtype, tagvalue))
                            pass
                        #
                    else:
                        val = tagvalue
                        if tagvalue != "":
                            found = True
                        #
                    #

                    if val is None:
                        val = ""
                        if opdata["rmblnk"] == 1:
                            opdata["rmblnk"] = 2
                        if opdata["rmblnkcol"] == 1:
                            opdata["rmblnkcol"] = 2
                    #

                    if found and val.find(":picture:") >= 0:
                        tagtype = s.TTYP_PIC
                        val = val.replace(":picture:", "")
                    #
                    # sqlmsg("Tag: {} Tagtype: {} Node:{} {} ".format(p_idx,tagtype,nodevalue, tagvalue[p_idx]))

                    if (
                        val is not None
                        and tagtype == s.TTYP_EMBEDHTML
                        and p_xmltext.find(tag) >= 0
                        or (tag.find("usr_htmlsig") >= 0 and p_xmltext.find(tag) >= 0)
                    ):
                        try:

                            localtree = LET.parse(io.StringIO(val), s.parser)
                            fakebody = localtree.find("body")
                            fakebody.tag = "span"
                            p_xmltext = p_xmltext.replace(tag, "")
                            p_xmlnode.append(fakebody)

                        except Exception as e:
                            # p_xmltext = p_xmltext.replace(tag, "Invalid User Signature {}".format(e))
                            sqlmsg(
                                "Invalid User Signature {}".format(e, tag),
                                p_type="local error",
                            )
                        #

                        # if p_xmlnode is not None:
                        #   customnode = LET.XML(val)
                        #   p_xmlnode.append(customnode)
                        #   p_xmltext = p_xmltext.replace(tag, "")
                        # else:
                        #   p_xmltext = p_xmltext.replace(tag, val)
                        # #
                    #
                    # if tagtype == s.TTYP_PIC: continue
                    if tagtype == s.TTYP_PIC:

                        if (
                            p_xmltext is not None
                            and p_xmltext.find(tag) >= 0
                            and s.HTML
                        ):
                            p_str = p_xmltext
                            # sqlmsg("Img Tag: {} Tagtype: {} Node:{} {} ".format(p_idx,tagtype,nodevalue, tagvalue[p_idx]))

                            def findtagdetails(p_wtype):
                                nonlocal p_str
                                lval = ""
                                typel = len(p_wtype)
                                i_ts = p_str.find(tag)
                                i_sts = p_str.find("[^", i_ts)
                                # sqlmsg("add [2] findtagdetails for tag:{} , i_ts:{}, i_sts:{},p_start:{},p_str:{},p_wtype:{}".format(tag,i_ts,i_sts,p_str[i_sts:],p_str,p_wtype), p_title="pl_mailmerge", p_type="local notice")
                                if i_sts >= 0 and i_sts >= i_ts:
                                    i_ste = p_str.find("^]", i_sts) + 2

                                    i_ws = p_str.find(
                                        p_wtype.lower() + "=", i_sts, i_ste + 2
                                    )
                                    if i_ws < 0:
                                        i_ws = p_str.find(
                                            p_wtype.upper() + "=", i_sts, i_ste
                                        )
                                    #
                                    i_we = p_str.find(",", i_ws, i_ste + 2)
                                    if i_we < 0:
                                        i_we = p_str.find("^", i_ws, i_ste)
                                    #

                                    if i_we >= 0 and i_ws >= 0:
                                        lval = p_str[i_ws + typel + 1 : i_we]
                                    #
                                    if lval != "":
                                        p_str = p_str[:i_sts] + p_str[i_ste:]
                                    #
                                    # sqlmsg("add [i] findtagdetails for tag:{} , i_sts:{}, i_ste:{},i_ws:{},i_we:{}  ,lval:{},p_str:{}".format(tag,i_sts,i_ste,i_ws,i_we,lval,p_str), p_title="pl_mailmerge", p_type="local notice")
                                #
                                if lval.isalnum():
                                    return float(lval)
                                else:
                                    return 0
                                #

                            #

                            h = float(fields.get(tag).get("h", "0.5")) * 96
                            # w = float(fields.get(tag).get("w", "2")) * 96
                            # h = float(0)
                            w = float(0)

                            # sqlmsg("test findtagdetails for tag:{} ,val:{}".format(tag,p_str), p_title="pl_mailmerge", p_type="local notice")
                            user_mw = findtagdetails("mw")
                            user_mh = findtagdetails("mh")
                            user_w = findtagdetails("w")
                            user_h = findtagdetails("h")
                            if user_w > 0 or user_h > 0:
                                w = user_w
                                h = user_h
                            #

                            p_xmltext = p_str

                            # width="{}%" height:"{}%"
                            style = ""

                            if user_mw > 0:
                                style += "max-width: {}px; ".format(user_mw)
                            if user_mh > 0:
                                style += "max-height: {}px; ".format(user_mh)

                            if w > 0 and user_mw <= 0:
                                style += "width: {}px; ".format(w)
                            if h > 0 and user_mh <= 0:
                                style += "height: {}px; ".format(h)
                            # sqlmsg("style for tag:{} ,style: {}, val:{}".format(tag,style,p_str), p_title="pl_mailmerge", p_type="local notice")

                            # sqlmsg("test findtagdetails for tag:{} , user_mw:{} , user_mh:{} ,user_w:{},user_h:{} style:{}".format(tag,user_mw,user_mh,user_w,user_h, style), p_title="pl_mailmerge", p_type="local notice")

                            imgstr = '<img alt="mergedimage" class="mergedimage nocid" src="data:image/png;base64,{}" style="{}"  />'.format(
                                val, style
                            )
                            if p_xmlnode is not None:
                                imgnode = LET.XML(imgstr)
                                p_xmlnode.append(imgnode)
                                p_xmltext = p_xmltext.replace(tag, "")
                                # sqlmsg("Img Got here with idx: {} img: {} val:{} {} ".format(p_idx,nodevalue, tagtype, tagvalue[p_idx]))
                            else:
                                p_xmltext = p_xmltext.replace(tag, imgstr)
                            #

                            p_xmltext = s.cleanImageTags(p_xmltext)

                        #
                        found = True
                        continue
                    #

                    if not found:
                        continue
                    #

                    if tagtype == s.TTYP_TBLROOT:
                        continue  # we do not handle inner values here.
                    if tagtype == s.TTYP_CONDITIONAL:
                        if p_xmltext is not None and p_xmltext.find(tag) >= 0:
                            if tryInt(tagvalue) > 0:
                                if tag.find("rmr!") >= 0:
                                    val = s.TAG_COND_RMR
                                    # sqlmsg("Will Remove Row: {} , {}, p_str={}".format(tag, tagvalue, nodevalue))
                                elif tag.find("rmc!") >= 0:
                                    val = s.TAG_COND_RMC
                                elif tag.find("cc!") >= 0:
                                    val = s.TAG_COND_CLR
                                elif tag.find("!rmp!") >= 0:
                                    val = s.TAG_COND_RMP
                                else:
                                    continue
                                #
                                # plpy.notice("Conditional Type({}) Tag Value: {} = {}  - opdata={}".format(p_type,tag,tagvalue,opdata))
                            else:
                                continue
                            #
                        #
                    #

                    if val is None:
                        val = ""
                        if opdata["rmblnk"] == 1:
                            opdata["rmblnk"] = 2
                        if opdata["rmblnkcol"] == 1:
                            opdata["rmblnkcol"] = 2
                    #
                    if type(val) is not str:
                        plpy.warning(
                            "Expected string type value. {}".format(type(val), val)
                        )
                        opdata["error"] = "Value not of type string"
                        break
                        # return opdata
                    #
                    if val == "" or val == "0":
                        if opdata["rmblnk"] == 1:
                            opdata["rmblnk"] = 2
                        if opdata["rmblnkcol"] == 1:
                            opdata["rmblnkcol"] = 2

                        found = False
                    #

                    if found:
                        if p_xmltext is not None and p_xmltext.find(tag) > -1:
                            if opdata["rmdr"] == 1:
                                opdata["rmdr"] = 2
                            #

                            opdata["mergedtags"].append(tag)
                            if opdata["value"] is None:
                                opdata["value"] = val
                            #
                        #
                    #

                    p_xmltext = p_xmltext.replace(tag, val)

                    opdata["xmltext"] = p_xmltext

                    # sqlmsg("Tag Replace tag: {} Val: {}".format(tag,val), p_type="local notice")
                    # plpy.notice("Tag replace: {}={}  obj:{}".format(tag, val, p_xmlnode.text))
                except Exception as e:
                    sqlmsg(
                        "Failed to replace tag value. tag: {} Error: {}".format(tag, e),
                        p_type="local error",
                    )
                #
                #
            #
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            sqlmsg(
                "LP: Failed to replace tag value. Idx: {} Error:{} Trace:{}".format(
                    p_idx, e, str(traceback.format_tb(exc_tb, 30))
                ),
                p_type="local error",
            )
        #
        return opdata

    #

    def tag_value_replace(
        s, p_xmlnode, p_idx=None, p_fields=None, p_type="", p_tblid="", p_colidx=0
    ):
        """Scans through html objects and replace from the given values."""
        opdata = {
            "error": None,
            "rmblnk": 0,
            "rmdr": 0,
            "mergedtags": [],
            "value": None,
            "rmblnkcol": 0,
            "limitrow": 0,
            "str": "",
            "xmlnode": p_xmlnode,
        }

        try:
            # sqlmsg("tag_value_replace: p_xmlnode={} p_xmlnode.tail={}".format(p_xmlnode,p_xmlnode.tail) , p_type="local error")
            if p_xmlnode.tail is not None and p_xmlnode.tail != "":
                o = s.tag_value_replace_node(
                    p_xmlnode,
                    p_xmlnode.tail,
                    p_idx,
                    p_fields,
                    p_type,
                    p_tblid,
                    p_colidx,
                )
                if o["xmltext"]:
                    p_xmlnode.tail = o["xmltext"]
                #
                opdata["str"] = p_xmlnode.tail
                # sqlmsg("Tail Replaced Tag={} Path={} Tail: {} Text={} o:{}".format(p_xmlnode.tag, s.htmltree.getpath(p_xmlnode),p_xmlnode.tail,opdata["str"], o))
            #
            # if p_xmlnode.head is not None:
            #   o = s.tag_value_replace_node(p_xmlnode, p_xmlnode.head, p_idx,p_fields,p_type,p_tblid, p_colidx)
            #   if o["xmltext"]:
            #     p_xmlnode.head = o["xmltext"]
            #   #
            #   opdata["str"] = p_xmlnode.head
            # #
            if p_xmlnode.text is not None:
                o = s.tag_value_replace_node(
                    p_xmlnode,
                    p_xmlnode.text,
                    p_idx,
                    p_fields,
                    p_type,
                    p_tblid,
                    p_colidx,
                )
                if o["xmltext"]:
                    p_xmlnode.text = o["xmltext"]
                #
                opdata["str"] = p_xmlnode.text
                # sqlmsg("Text Replaced Tag={} Path={} Text={} o:{}".format(p_xmlnode.tag, s.htmltree.getpath(p_xmlnode),opdata["str"], o))
            #
            elif type(p_xmlnode) == "str":
                o = s.tag_value_replace_node(
                    None, p_xmlnode, p_idx, p_fields, p_type, p_tblid, p_colidx
                )
                if o["xmltext"]:
                    p_xmlnode = o["xmltext"]
                #
                opdata["str"] = p_xmlnode
            #
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            sqlmsg(
                "LP: Failed to replace tag value. Idx: {} Error:{} Trace:{}".format(
                    p_idx, e, str(traceback.format_tb(exc_tb, 30))
                ),
                p_type="local error",
            )
        #
        return opdata

    #

    def looptables(s, parent):
        for elem in parent:
            if elem.tag == "table":
                yield elem
                continue
            #
            for inner in s.looptables(elem):
                yield inner
            #
        #

    #

    def looprow(s, parent):
        for elem in parent:
            if elem.tag == "tr":
                yield elem
                continue
            #
            for inner in s.looprow(elem):
                yield inner
            #
        #

    #

    def loopcell(s, parent):
        for elem in parent:
            if elem.tag == "td":
                yield elem
                continue
            #
            for inner in s.loopcell(elem):
                yield inner
            #
        #

    #
    def tbl_post_loop(s, p_tables):
        pass

    #

    def tbl_basic_replace(s, p_tables, p_parent=0):
        if p_parent == 0:
            s.tbl_basic_replace_cnt = 0

        for tag in p_tables:
            if tag.tag == "table":
                s.tbl_basic_replace_cnt += 1
                tblsource = "{},{}".format(s.tbl_basic_replace_cnt, p_parent)
                tableName = tblsource
                if p_parent > 0:
                    if s.tbl_map.get(str(p_parent)) is None:
                        s.tbl_map[str(p_parent)] = {
                            "tblptr": tag,
                            "children": None,
                            "nr": tblsource,
                            "name": tableName,
                        }
                    #
                    if s.tbl_map[str(p_parent)]["children"] is None:
                        s.tbl_map[str(p_parent)]["children"] = {}
                    #
                    s.tbl_map[str(p_parent)]["children"][
                        str(s.tbl_basic_replace_cnt)
                    ] = {
                        "tblptr": tag,
                        "children": None,
                        "nr": tblsource,
                        "name": tableName,
                    }
                else:
                    s.tbl_map[str(s.tbl_basic_replace_cnt)] = {
                        "tblptr": tag,
                        "children": None,
                        "nr": tblsource,
                        "name": tableName,
                    }
                #

                lastid = s.tbl_basic_replace_cnt
                distinctdata = []
                rowcnt = -1
                for row in s.looprow(tag):
                    colidx = 0
                    rowcnt += 1
                    rmrow = False
                    for cell in s.loopcell(row):
                        rmcol = False
                        if cell is None or cell.text is None:
                            continue
                        opdata = {}
                        for subtag in cell.iter():
                            # plpy.notice("Tag: {} Path: {} Text:{}".format(tag.tag, htmltree.getpath(tag),tag.text))
                            if subtag.text and subtag.text != "":
                                # dat = s.str_value_replace(subtag.text, None, s.inputdata.get('fields'),p_htmlnode=subtag)
                                opdata = s.tag_value_replace(
                                    subtag,
                                    None,
                                    None,
                                    "tablebasic",
                                    str(lastid),
                                    colidx,
                                )
                                # subtag.text = dat["str"]
                                # subtag.text = s.eval_replace(dat["str"], "html")
                            #
                            if subtag.tail and subtag.tail != "":
                                # dat = s.str_value_replace(subtag.text, None, s.inputdata.get('fields'),p_htmlnode=subtag)
                                opdata = s.tag_value_replace(
                                    subtag,
                                    None,
                                    None,
                                    "tablebasic",
                                    str(lastid),
                                    colidx,
                                )
                                # subtag.text = dat["str"]
                                # subtag.text = s.eval_replace(dat["str"], "html")
                            #

                        #

                        if cell.text.find(s.STR_TAG_FREETBL) > 0:
                            s.tbl_mode[str(p_parent)] = "free"
                        #
                        # remove the blank value row again. Note that its not the filter and depends on value of field with this remove tag.
                        if opdata.get("rmblnk", 0) == 2:
                            rmrow = True
                        if opdata.get("rmblnkcol", 0) == 2:
                            rmcol = True

                        if opdata.get("rmdr", 0) == 2:
                            rmrow = True
                        if (
                            len(opdata.get("mergedtags", [])) > 0
                        ):  # this will reset the tag, important to detect sibling
                            rmrow = False
                            rmcol = False
                        #

                        if rmcol:
                            s.tbl_rmcol.append(cell)
                        #
                        colidx += 1
                    #
                    if rmrow:
                        s.tbl_rmrow.append(row)
                    #
                #

            #
        #

    #

    def tbl_complex_init(s):
        if s.inputdata is None:
            return
        if s.tbl_map is None:
            return
        complexfields = s.inputdata.get("complexfields")
        if complexfields is None:
            return

        for nm in s.tbl_map:
            tbldat = s.tbl_map.get(nm)
            if tbldat is None:
                continue
            srctbl = tbldat.get("tblptr")
            tblfields = complexfields.get(nm)
            if tblfields is None or srctbl is None:
                plpy.notice(
                    "Table or data [{}] not found for fields -> {}".format(
                        nm, tblfields
                    )
                )
                continue
            #

            # plpy.notice("tbl_complex_init: ", srctbl, tbldat)

            s.tbl_complex_copy(nm, srctbl, tblfields)
        #

        for r in s.tbl_rmrow:
            # sqlmsg("Removing row: len: {} -> {}".format(len(s.tbl_rmrow), r) )
            parent = r._tr.find("..")
            if parent is not None:
                parent.remove(r._tr)
                # sqlmsg("Removed row: {}".format(r))
            #
        #

    #

    def tbl_complex_copy(s, p_tblid, p_tbl, p_fields):
        # plpy.notice("TBLID: {}  ,dat\r\n{}".format(p_tblid,p_fields))
        mappedtbl = s.tbl_map.get(p_tblid)
        children = None
        distinctdata = {}
        limitrow = 0
        fieldtags = list(p_fields.keys())
        if mappedtbl is not None:
            children = mappedtbl.get("children")
            limitrow = mappedtbl.get("limitrow", 0)
        #

        # plpy.notice("tbl_complex_copy : ", p_tblid,p_tbl, p_fields.keys() )

        # if p_fields is None: return
        if p_tbl is None:
            return

        dat = None
        datcnt = 0
        for t in p_fields.keys():
            fielval = p_fields.get(t)
            # plpy.notice("tbl_complex_copy field: datcnt ", datcnt, t)
            if fielval is None:
                continue
            dat = fielval.get("value", None)
            if dat is None:
                continue
            if isinstance(dat, dict) or isinstance(dat, list) or isinstance(dat, tuple):
                datcnt = len(dat)
            else:
                continue
            #
            break
        #

        if children is not None:
            for tblid in children:
                childfields = p_fields.get(tblid, {}).get("value", None)
                if childfields is not None and type(childfields) is list:
                    # plpy.notice("Childfields: {}".format(list(childfields[0].keys())))
                    fieldtags.extend(list(childfields[0].keys()))
                #
            #
        #
        if s.tbl_complex_rowdat.get(p_tblid) is None:
            s.tbl_complex_rowdat[p_tblid] = {"min": 0, "max": 0}
        #
        s.tbl_complex_rowdat[p_tblid]["max"] = datcnt

        rowdata = s.tbl_complex_headerstart(p_tbl, fieldtags, p_tblid, p_fields)
        # plpy.notice("TBLID: {} len:{},dat\r\n{}\r\nRowdata:{}".format(p_tblid, datcnt,p_fields,rowdata))
        lastrow = None
        for rn in range(
            0, datcnt
        ):  # loop through table data which is the rows for the first column.
            if limitrow > 0 and rn >= limitrow:
                break
            #
            for (
                rd
            ) in (
                rowdata
            ):  ##this will usually be 1 for 1 row to copy. but can be more than 1 (Generate Blank row feature)
                # plpy.notice("Copies: {} rd:{}".format(rn, rd))

                if not rowdata[rd]["merged"]:
                    if rowdata[rd].get("type", 0) == 4 and rn != datcnt - 1:
                        continue
                    #
                    lastrow = rowdata[rd]["row"]
                    rowdata[rd]["merged"] = True

                else:
                    newrow = copy.deepcopy(rowdata[rd]["rowcopy"])
                    lasttr = lastrow.addnext(newrow)
                    lastrow = newrow
                #

                rmrow = False
                for cell in s.loopcell(lastrow):
                    colidx = 0
                    rmcol = False

                    opdata = s.tag_value_replaceAll(
                        cell, rn, p_fields, "table", p_tblid, colidx
                    )

                    if cell.text.find(s.TAG_COND_TBL_DISTINCT) >= 0:
                        if distinctdata.get(str(rn), None) is None:
                            distinctdata[str(rn)] = {}
                        distinctdata[str(rn)][str(colidx)] = opdata.get("value", "")

                        # sqlmsg("Test distinct: {}".format(distinctdata))
                        for ir in distinctdata:
                            if int(ir) == int(rn):
                                continue
                            prev = distinctdata.get(str(ir), {}).get(str(colidx), "")
                            if (
                                prev != ""
                                and prev == distinctdata[str(rn)][str(colidx)]
                            ):
                                rmrow = True
                                # sqlmsg("Found distinct: col:{} row:{},{} val:{} rmrow={} mergedtags={} \r\nText={}".format(colidx,rn,ir,prev, rmrow,opdata.get("mergedtags", []), run.text))
                                break
                            #
                        #

                    #
                    # remove the blank value row again. Note that its not the filter and depends on value of field with this remove tag.
                    if opdata.get("rmblnk", 0) == 2:
                        rmrow = True
                    if opdata.get("rmblnkcol", 0) == 2:
                        rmcol = True
                    if opdata.get("rmdr", 0) == 2:
                        rmrow = True

                    ############################################################################################################
                    celltable = s.looptables(cell)
                    if (
                        children is not None
                        and celltable is not None
                        and len(celltable) > 0
                    ):
                        for tblid in children:
                            # plpy.notice("Table in cell: Row {} Tbl nr: {} tblid: {}".format(rn,tblnr, tblid))
                            for tblptr in celltable:
                                childfields = p_fields.get(tblid, {}).get("value", None)
                                if (
                                    childfields is not None
                                    and type(childfields) is list
                                    and type(childfields[rn]) is dict
                                ):
                                    # plpy.notice("Found child table: {}  in cell {}".format(tblid, tblptr))
                                    s.tbl_complex_copy(tblid, tblptr, childfields[rn])
                                #

                                # plpy.notice("Child table found: {} {}".format(tblid, tblptr))
                            #
                        #
                    #
                    ############################################################################################################

                    if s.tbl_mode.get(p_tblid) == "free":
                        for tbl in celltable:
                            for row2 in s.looprow(tbl):
                                for cell2 in s.loopcell(row2):
                                    s.tag_value_replaceAll(
                                        cell2, rn, p_fields, "table", p_tblid, colidx
                                    )
                                #
                            #
                        #
                    #

                    if rmcol:
                        s.tbl_rmcol.append(cell)
                    #

                    colidx += 1
                #
                # sqlmsg("row  -> row:{} rm:{}".format(rn,rmrow))
                if rmrow:
                    s.tbl_rmrow.append(lastrow)
                #
            #
        #

    #

    def tbl_complex_headerstart(s, p_tbl, p_taglist, p_tblid, p_fields):
        rownr = 0
        templaterows = {}

        if s.tbl_complex_rowdat.get(p_tblid) is None:
            s.tbl_complex_rowdat[p_tblid] = {"min": 0, "max": 0}
        #

        for row in s.looprow(p_tbl):
            breakrow = False
            for cell in s.loopcell(row):
                if cell is None or cell.text is None:
                    continue
                bodytext = cell.text
                # check fields in paragraph
                for fieldname in p_taglist:
                    if type(fieldname) is dict:
                        continue  # we dont want the sub table data here.
                    if bodytext.find(fieldname) >= 0:
                        # if int(p_fields.get(fieldname,{}).get("type", 0)) == 4:
                        #  continue
                        #
                        templaterows[rownr] = {
                            "row": row,
                            "rowcopy": copy.deepcopy(row),
                            "merged": False,
                            "type": int(p_fields.get(fieldname, {}).get("type", 0)),
                        }
                        rownr += 1
                        breakrow = True
                        if s.tbl_complex_rowdat[p_tblid]["min"] <= 0:
                            s.tbl_complex_rowdat[p_tblid]["min"] = rownr
                        #
                        break
                    #
                    # handle blank row
                    elif bodytext.find(s.TAG_S + "blnkrow" + s.TAG_E) >= 0:
                        templaterows[rownr] = {
                            "row": row,
                            "rowcopy": copy.deepcopy(row),
                            "merged": False,
                            "type": 0,
                        }
                        rownr += 1
                        breakrow = True
                        if s.tbl_complex_rowdat[p_tblid]["min"] <= 0:
                            s.tbl_complex_rowdat[p_tblid]["min"] = rownr
                        #
                        break
                    #
                #
                if breakrow:
                    break
                for childtable in s.looptables(cell):
                    for childrow in s.looprow(childtable):
                        for childcell in s.loopcell(childrow):
                            bodytext = childcell.text
                            for fieldname in p_taglist:
                                if type(fieldname) is dict:
                                    continue  # we dont want the sub table data here.
                                if bodytext.find(fieldname) >= 0:
                                    templaterows[rownr] = {
                                        "row": row,
                                        "rowcopy": copy.deepcopy(row),
                                        "merged": False,
                                        "type": 0,
                                    }
                                    rownr += 1
                                    breakrow = True
                                    break
                                #
                            #
                            if breakrow:
                                break
                        #
                        if breakrow:
                            break
                    #
                    if breakrow:
                        break
                #
                if breakrow:
                    break
            #
        #
        return templaterows

    #

    def mark_html(s):
        global applyError, errorstr
        applyError = False
        errorstr = ""
        content = ""

        if len(s.inputdata) <= 0:
            raise MergeToolWarning("No tag data given.")
        #

        errortags = s.inputdata.get("marktags")
        if errortags is None:
            raise Exception(
                "Expected an array list of tags. Got: {}".format(s.inputdata)
            )
        #

        with open(s.srcfilename, "rb") as f:
            data = f.read()
            content = data.decode("utf8", "replace")
        #

        def search_str(p_str, p_tag_s, p_tag_e, p_start=0):
            ###Inner Routine to search and add faults
            global applyError, errorstr
            slen = len(p_str)
            i_s = p_str.lower().find(p_tag_s, p_start)
            i_e = p_str.lower().find(p_tag_e, i_s)

            if i_s >= 0 and p_str[i_s + 1 : i_s + 1 + len(s.TAG_ERR)] == s.TAG_ERR:
                fulltag = p_str[i_s : i_e + len(p_tag_e)]
                s.add_fault("not found", fulltag)

                if s.DEBUG:
                    plpy.notice(
                        "Already error checked. {}".format(
                            p_str[i_s : i_e + len(p_tag_e)]
                        )
                    )
                #
                return p_str
            elif i_s >= 0 and i_e > i_s:
                fulltag = p_str[i_s : i_e + len(p_tag_e)]
                found = False
                if i_e - i_s > 100:
                    applyError = True
                    s.add_fault("Tag to long. More than 100 characters.", fulltag)
                #
                if p_values is not None:
                    for v in p_values:
                        errorstr = v.get("error", "Not found")
                        tblsrc = v.get("source", "")

                        if (
                            fulltag != ""
                            and v.get("tag", "").find(fulltag) >= 0
                            and (
                                tblsrc == p_tblsource
                                or p_tblsource == ""
                                or tblsrc == ""
                            )
                        ):
                            found = True
                            # plpy.notice("Found - {} -> {} Src: {}:{}".format(fulltag, v, p_tblsource, tblsrc))
                        #
                    #
                    if found:
                        # p_str = p_str[:i_s+len(p_tag_s)] + s.TAG_ERR + p_str[i_s+len(p_tag_s):]
                        p_str = p_str  # we are not replacing it anymore. Just add error below.
                        applyError = True
                        # s.add_fault("not found", fulltag)
                    #
                # plpy.notice("Cheking tag: {} found: {} Started@{}".format(fulltag,found,p_start))
                p_str = search_str(
                    p_str, p_tag_s, p_tag_e, i_e + len(p_tag_e) + (len(p_str) - slen)
                )
            elif i_s >= 0:
                if s.DEBUG:
                    plpy.notice(
                        "Broken tag: {} Started@{}".format(
                            p_str[i_s + len(p_tag_s) : i_s + len(p_tag_s) + 30], p_start
                        )
                    )
                #
                p_str = (
                    p_str[: i_s + len(p_tag_s)]
                    + s.TAG_ERR
                    + p_str[i_s + len(p_tag_s) :]
                )
                p_str = search_str(
                    p_str,
                    p_tag_s,
                    p_tag_e,
                    i_s + len(p_tag_e) + len(p_tag_s) + (len(p_str) - slen),
                )  # search after, start + tags sizes + diff in inserted text.
                applyError = True
                s.add_fault("incomplete tag", "incomplete " + p_tag_s + "")
            #
            return p_str

        ###

        search_str(content, s.TAG_S, s.TAG_E)

        with open(s.destfilename, "wb") as f:
            data = f.write(content.encode("utf8"))
        #

        if sys.platform.lower() == "linux":
            os.chmod(s.destfilename, 0o777)
        #

    #

    def convert_html(s):
        replacekeys = {}
        content = ""
        with open(s.srcfilename, "rb") as f:
            data = f.read()
            content = data.decode("utf8", "replace")
        #

        tagdata = s.inputdata.get("tags")
        if s.inputdata is None or len(s.inputdata) <= 0:
            raise MergeToolWarning("No tag data given.")
        #
        replacekeys = s.inputdata.get(
            "replacetags"
        )  ##expected {"replacetags": {"key1":"newkey1", "key2":"newkey2"}
        content = s.replace_str(content, replacekeys)

        with open(s.destfilename, "wb") as f:
            data = f.write(content.encode("utf8"))
        #
        if sys.platform.lower() == "linux":
            os.chmod(s.destfilename, 0o777)
        #

    #

    def generate_mockup(s):
        content = """<html>
    <head>
    <meta charset="UTF8">
    </head>
    <body>
    """
        lastsection = ""
        lasttable = ""

        for tag in s.inputdata:
            if type(s.inputdata.get(tag)) is not dict:
                raise MergeToolError(
                    "Expecting tag field to have a type and value property.",
                    "tag_value_replace",
                )
            #
            tagtype = int(s.inputdata.get(tag).get("type", 0))
            tooltip = s.inputdata.get(tag).get("tooltip", "")
            display_name = s.inputdata.get(tag).get("display_name", "")
            parent_name = s.inputdata.get(tag).get("parent_name", "")
            parent_tooltip = s.inputdata.get(tag).get("parent_tooltip", "")
            parent_table_name = s.inputdata.get(tag).get("parent_table", "")
            table_name = s.inputdata.get(tag).get("table_name", "")
            if table_name == "":
                table_name = parent_table_name

            if parent_name != lastsection:
                lastsection = parent_name
                content += "<p><b>{}:</b>".format(parent_name)
                content += "<br />Tooltip:{}<br />".format(tooltip)
                content += "</p>"
            #
            if tagtype in (0, 1, 3):
                content += "<p><b>{}:</b><br />Tag:{}".format(
                    display_name, tag.strip(s.TAG_S).strip(s.TAG_E)
                )
                content += "<br />Tooltip:{}<br />".format(tooltip)
                content += "<br />Data:[{}]<br />".format(tag)
                content += "</p>"
            elif tagtype == 2:
                if lasttable != parent_name:
                    lasttable = parent_name
                    plpy.notice(
                        "new table: {}, {}, {}".format(parent_name, display_name, tag)
                    )
                    # if objtable
                #
                plpy.notice(
                    "last table: {}, {}, {}, {}".format(
                        table_name, parent_table_name, display_name, tag
                    )
                )
            #
        #

        content += "</body></html>"
        with open(s.destfilename, "wb") as f:
            data = f.write(content.encode("utf8"))
        #
        if sys.platform.lower() == "linux":
            os.chmod(s.destfilename, 0o777)
        #

    #


#
