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


class MergeToolTEXT(MergeTool):
    """Bitech TEXT Merge Tool"""

    def __init__(s):
        super().__init__()

        s.srcdoc = None
        s.destdoc = None
        s.srcfilename = None
        s.destfilename = None
        s.visit_tables_cnt = 0
        s.tbl_basic_replace_cnt = 0
        s.tbl_map = {}
        s.tbl_mode = {}

    def execute(s):
        if s.mode.find("peek") >= 0:
            tagdata = s.peek()
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
            s.merge()
        elif s.mode.find("mocktest") >= 0:
            s.generate_mockup()
        elif s.mode.find("mark") >= 0:
            s.mark()

        else:
            raise MergeToolError(
                "Mode {} not supported.".format(s.mode), "MergeToolHTML.init"
            )
        #

        return s.results

    def peek(s):
        """Scans through a string for merge tags and returns a list of tags found within merge tags."""
        tags = {}
        tags["doctags"] = []
        tags["pictures"] = []
        tags["headertags"] = []
        tags["footertags"] = []
        tags["tabletags"] = []
        content = s.load_generic()

        tabledata = s.visit_table_tags_peek(content, s.inputdata.get("fields"))
        tags["tabletags"] = tabledata["tags"]
        content = tabledata["str"]

        tags["doctags"] = s.peek_str(content)

        return tags

    #

    def merge(s):
        content = s.load_generic()

        dat = s.str_value_replace(content, None, s.inputdata.get("fields"))

        content = dat["str"]
        content = s.visit_table_tags(content, s.inputdata.get("fields"))
        content = s.eval_replace(dat["str"], "txt")
        s.save_generic(content)

    #
    def mark(s):
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

        content = s.load_generic()

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
        s.save_generic(content)

    #

    def convert(s):
        replacekeys = {}
        content = s.load_generic()

        tagdata = s.inputdata.get("tags")
        if s.inputdata is None or len(s.inputdata) <= 0:
            raise MergeToolWarning("No tag data given.")
        #
        replacekeys = s.inputdata.get(
            "replacetags"
        )  ##expected {"replacetags": {"key1":"newkey1", "key2":"newkey2"}
        content = s.replace_str(content, replacekeys)

        s.save_generic(content)

    #

    def generate_mockup(s):
        content = """

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
                content += "\r\n{}:\r\n".format(parent_name)
                content += "\r\nTooltip:{}\r\n".format(tooltip)
                content += "\r\n"
            #
            if tagtype in (0, 1, 3):
                content += "\r\n\r\n{}:\r\nTag:{}".format(
                    display_name, tag.strip(s.TAG_S).strip(s.TAG_E)
                )
                content += "\r\nTooltip:{}".format(tooltip)
                content += "\r\nData:[{}]".format(tag)
                content += "\r\n"
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

        content += ""
        s.save_generic(content)

    #

    def visit_table_tags(s, p_str, p_fields=None, p_parent=0):
        result = {}
        result["tags"] = {}
        result["str"] = p_str
        complexfields = s.inputdata.get("complexfields")

        # sqlmsg("visit_table_tags attempt: {} t ".format(p_str))

        openindex = p_str.find(s.TAG_S_TBLM)
        if openindex < 0:
            return result
        #
        closeindex = s.getNextCloseTableTagPos(p_str, openindex + len(s.TAG_S_TBLM))
        if not (closeindex > openindex and closeindex >= 0):
            return result
        #
        closeindex = closeindex + len(s.TAG_E_TBLM)
        tmpstr = ""
        s.visit_tables_cnt += 1
        tableid = s.visit_tables_cnt
        tablelookup = "{},{}".format(tableid, p_parent)

        if result["tags"].get(tablelookup) is None:
            result["tags"][tablelookup] = []
        #

        tableinner = p_str[
            openindex + len(s.TAG_S_TBLM) : closeindex - len(s.TAG_E_TBLM)
        ]
        # sqlmsg("visit_table_tags tablelookup: {} tableinner:{} ".format(tablelookup, tableinner))
        for i in range(tableinner.count(s.TAG_S_TBLM)):
            res = s.visit_table_tags(tableinner, p_fields, tableid)
            result["tags"].update(res["tags"])
            tableinner = res["str"]
        #
        result["tags"][tablelookup].extend(s.peek_str(tableinner, False))

        tblfields = complexfields.get(tablelookup)
        fieldtags = []
        datcnt = 0
        # sqlmsg("visit_table_tags tablelookup: {} tblfields:{} ".format(tablelookup,tblfields))
        if tblfields is not None:
            fieldtags = list(tblfields.keys())
            fielval = tblfields.get(fieldtags[0])
            if fielval is not None:
                dat = fielval.get("value", None)
                if (
                    isinstance(dat, dict)
                    or isinstance(dat, list)
                    or isinstance(dat, tuple)
                ):
                    datcnt = len(dat)
                #
            #
        #
        for ci in result["tags"][tablelookup]:
            tag = result["tags"][tablelookup][ci]
            rowagg = ""
            for rn in range(0, datcnt):
                dat = s.str_value_replace(tableinner, rn, p_fields, "", tablelookup, 0)
                rowagg = "{},{}".format(rowagg, dat["str"])
            #
            tableinner = tableinner.replace(tag, rowagg)
        #

        dat = s.str_value_replace(tableinner, None, p_fields)
        tableinner = dat["str"]
        tableinner = s.eval_replace(tableinner, "txt")

        # sqlmsg("Visit table peek : {} result:{}".format(p_str[openindex:closeindex], result))
        p_str = p_str[:openindex] + tableinner + p_str[closeindex:]

        if p_str.find(s.TAG_S_TBLM):
            res = s.visit_table_tags(p_str, p_fields, 0)
            result["tags"].update(res["tags"])
            p_str = res["str"]
        #

        result["str"] = p_str

        return result

    #

    def getNextCloseTableTagPos(s, p_str, p_start):
        cnt = 0
        openindex = p_start
        slen = len(s.TAG_S_TBLM)
        elen = len(s.TAG_E_TBLM)
        for i in range(p_start + slen, len(p_str)):
            if p_str[i : i + slen] == s.TAG_S_TBLM:
                openindex = i + slen
                cnt += 1
            #
            if p_str[i : i + slen] == s.TAG_E_TBLM:
                break
            #
        #

        # sqlmsg("getNextTableTagPos openindex:{} nextend:{} cnt:{} skipcnt:{} str:{}".format(openindex,nextend,cnt,skipcnt, p_str ))
        nextend = p_str.find(s.TAG_E_TBLM)
        for i in range(cnt):
            closeindex = p_str.find(s.TAG_E_TBLM, nextend + len(s.TAG_E_TBLM))
            # sqlmsg("getNextTableTagPos i_{} cnt_{},{} p_start={} nextpos:{} closeindex:{} str:{}".format(i,cnt,skipcnt, p_start, nextpos,closeindex,p_str ))
            if closeindex < 0:
                break
            #
            nextend = closeindex
        #

        # sqlmsg("getNextTableTagPos openindex:{} nextend:{} cnt:{}str:{}".format(openindex,nextend,cnt, p_str ))

        return nextend

    #

    def visit_table_tags_peek(s, p_str, p_fields=None, p_parent=0):
        result = {}
        result["tags"] = {}
        result["str"] = p_str

        openindex = p_str.find(s.TAG_S_TBLM)
        if openindex < 0:
            return result
        #
        closeindex = s.getNextCloseTableTagPos(p_str, openindex + len(s.TAG_S_TBLM))
        if not (closeindex > openindex and closeindex >= 0):
            return result
        #
        closeindex = closeindex + len(s.TAG_E_TBLM)
        tmpstr = ""
        s.visit_tables_cnt += 1
        tableid = s.visit_tables_cnt

        if result["tags"].get("{},{}".format(tableid, p_parent)) is None:
            result["tags"]["{},{}".format(tableid, p_parent)] = []
        #

        tableinner = p_str[
            openindex + len(s.TAG_S_TBLM) : closeindex - len(s.TAG_E_TBLM)
        ]
        for i in range(tableinner.count(s.TAG_S_TBLM)):
            res = s.visit_table_tags_peek(tableinner, p_fields, tableid)
            result["tags"].update(res["tags"])
            tableinner = res["str"]
        #

        result["tags"]["{},{}".format(tableid, p_parent)].extend(
            s.peek_str(tableinner, False)
        )

        # sqlmsg("Visit table peek : {} result:{}".format(p_str[openindex:closeindex], result))
        p_str = p_str[:openindex] + tmpstr + p_str[closeindex:]

        if p_str.find(s.TAG_S_TBLM):
            res = s.visit_table_tags_peek(p_str, p_fields, 0)
            result["tags"].update(res["tags"])
            p_str = res["str"]
        #

        result["str"] = p_str

        return result

    #
