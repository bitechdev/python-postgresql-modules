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


class MergeTool(object):
    """Bitech Merge Tool"""

    DEBUG = False
    TAG_S_SP = chr(171)
    TAG_E_SP = chr(187)
    TAG_S = "[*"
    TAG_E = "*]"

    TAG_S_OLD = "[~"
    TAG_E_OLD = "~]"

    TAG_S_EVAL = "[="
    TAG_E_EVAL = "=]"

    TAG_S_INPLACE = "[I="
    TAG_E_INPLACE = "=I]"

    TAG_S_ALIASDECL = "[A="
    TAG_E_ALIASDECL = "=A]"
    TAG_D_ALIASDECL = "="

    TAG_S_ALIASUSE = "[A_"
    TAG_E_ALIASUSE = "_A]"

    TAG_S_POSTOP = "[+"
    TAG_E_POSTOP = "+]"

    TAG_S_TBLM = "[T_"
    TAG_E_TBLM = "_T]"

    TAG_ERR = "!ERROR!"
    STR_ERR = "!Fix tag!"
    TAG_SPLT = "|"
    TAG_RULESPLT = ":"
    STR_TAG_RMBD = "rmbd!row"
    STR_TAG_RMBC = "rmbd!col"
    STR_TAG_RMDR = "rmdr!tbl"

    STR_TAG_FREETBL = "free!tbl"
    STR_TAG_ESIGN = "!esign"
    STR_TAG_LIMITROW = "limitrow"
    STR_TAG_TBLSEARCH_RM = "tbl!rm!search"

    TTYP_ROOT = 0
    TTYP_FIELD = 1
    TTYP_TBLFIELD = 2
    TTYP_TBLROOT = 3
    TTYP_AGGFIELD = 4
    TTYP_PIC = 5
    TTYP_SPECIAL = 6
    TTYP_FILTER = 7
    TTYP_CONDITIONAL = 8
    TTYP_DOCREPLACE = 9
    TTYP_EMBEDHTML = 10

    OP_TAGS = [
        "ordinal",
        "cardinal",
        "ordinalw",
        "cardinalw",
        "minvalue",
        "maxvalue",
        "minvaluew",
        "maxvaluew",
    ]
    CELL_OP_TAGS = ["col.min", "col.max", "col.avg", "col.sum", "col.count"]

    TAG_COND_RMR = "!!!rmr!!!"
    TAG_COND_RMC = "!!!rmc!!!"
    TAG_COND_CLR = "!!<clr>!!"
    TAG_COND_RTBL = "rm!tbl"
    TAG_COND_TBL_DISTINCT = "dtc!tbl"
    TAG_COND_RMP = "!!!rmp!!!"

    def __init__(s):
        s.mode = ""
        s.inputdata = None
        s.faults = []
        s.HTML = False
        s.cleanMissing = True  # this will remove the op tags and those not found.
        s.aliasList = []
        s.aliasListPeek = []
        s.template_blob = None
        s.srcfilename = None
        s.destfilename = None
        s.post_format = []
        s.results = None
        s.table_removerow_list = []
        s.tbl_col_operation = {}
        s.tagdebug = s.DEBUG
        s.visit_tables_uniq = []

    #

    def execute(s):
        return s.results

    def unique_table_check(s, ptr):
        for p in s.visit_tables_uniq:
            if p == ptr:
                return False
            #
        #
        s.visit_tables_uniq.append(ptr)
        return True

    #

    def save_generic(s, p_content):
        if s.destfilename is None or len(s.destfilename) < 3:
            s.results = p_content.encode("utf8")
        else:
            with open(s.destfilename, "wb") as f:
                data = f.write(p_content.encode("utf8"))
            #
            if sys.platform.lower() == "linux":
                os.chmod(s.destfilename, 0o777)
            #
        #
        return s.results

    #

    def load_generic(s):
        content = None
        if s.srcfilename is None or len(s.srcfilename) < 3:
            content = s.template_blob.decode("utf8", "replace")
        else:
            with open(s.srcfilename, "rb") as f:
                data = f.read()
                content = data.decode("utf8", "replace")
        #
        #
        return content

    #

    def add_fault(s, p_type, p_msg):
        s.faults.append({"tag": p_msg, "type": p_type})

    #

    def peek_str(s, p_str, p_inc_old=True):
        """Scans through a string for merge tags and returns a list of tags found within merge tags."""
        tags = []
        if p_str is None or p_str == "":
            return p_str

        def search_str(p_str, p_tag_s, p_tag_e, p_start=0, p_tagsrc="tags"):
            # global tags
            i_s = p_str.find(p_tag_s, p_start)
            i_e = p_str.find(p_tag_e, i_s)
            if i_s >= 0 and i_e > i_s:
                keyname = p_str[i_s + len(p_tag_s) : i_e]
                if len(keyname) > 1:
                    tags.append(p_tag_s + keyname + p_tag_e)
                    # recursive to next position.
                    search_str(p_str, p_tag_s, p_tag_e, i_e + len(p_tag_e), p_tagsrc)
                #
            #
            return p_str

        #

        search_str(p_str, s.TAG_S, s.TAG_E)
        if p_inc_old:
            search_str(p_str, s.TAG_S_OLD, s.TAG_E_OLD)
            search_str(p_str, s.TAG_S_SP, s.TAG_E_SP)
        #

        return tags

    #

    def buidAliasMap(s):
        """
        Rebuild alias maps from declare to use
        """
        # sqlmsg("Alias Map1: {}".format(s.aliasList), p_title="local error")
        records = len(s.aliasList)
        for k in range(records):
            if s.aliasList[k]["type"] != "use":
                continue  # only use
            if s.aliasList[k]["declid"] is not None:
                continue
            opFound = False
            ### These are placement operators
            for opv in s.aliasList[k]["ops"]:
                try:
                    if opv == "max":
                        lastdat = 0
                        decid = None
                        for l in range(records):
                            if s.aliasList[l]["type"] != "declare":
                                continue  # only declare
                            if s.aliasList[k]["name"] == s.aliasList[l]["name"]:

                                if s.aliasList[l]["datindex"] > lastdat:
                                    lastdat = s.aliasList[l]["datindex"]
                                    decid = l
                                #
                            #
                        #
                        if decid is not None:
                            s.aliasList[k]["declid"] = decid
                            opFound = True
                            continue
                        #
                    #
                    elif opv == "min":
                        lastdat = 0
                        decid = None
                        for l in range(records):
                            if s.aliasList[l]["type"] != "declare":
                                continue  # only declare
                            if s.aliasList[k]["name"] == s.aliasList[l]["name"]:

                                if s.aliasList[l]["datindex"] < lastdat:
                                    lastdat = s.aliasList[l]["datindex"]
                                    decid = l
                                #
                            #
                        #
                        if decid is not None:
                            s.aliasList[k]["declid"] = decid
                            opFound = True
                            continue
                        #
                    #
                    elif opv.find("n") == 0:
                        num = opv[1:]
                        decid = None
                        for l in range(records):
                            if s.aliasList[l]["type"] != "declare":
                                continue  # only declare
                            if s.aliasList[k]["name"] == s.aliasList[l]["name"]:
                                if s.aliasList[l]["datindex"] == num:
                                    decid = l
                                #
                            #
                        #
                        if decid is not None:
                            s.aliasList[k]["declid"] = decid
                            opFound = True
                            continue
                        #
                    #
                #
                except Exception as e:
                    sqlmsg(
                        "Alias Error: {} in : {}".format(e, s.aliasList[k]),
                        p_title="local notice",
                    )
                #
            if opFound:
                continue
            #

            ##fallback to default map
            for l in range(records):
                if s.aliasList[l]["type"] != "declare":
                    continue  # only declare
                if s.aliasList[k]["name"] == s.aliasList[l]["name"]:
                    s.aliasList[k]["declid"] = l
                    continue
                #
            #
        #
        # sqlmsg("Alias Map2: {}".format(s.aliasList), p_title="local error")

    #

    def alias_replace(s, pStr, pKey=None, pIndex=None):
        # sqlmsg("alias_replace aliasList={}".format(s.aliasList), "alias_replace")
        s.buidAliasMap()
        for alias in s.aliasList:
            if alias["type"] != "use":
                continue
            # sqlmsg(" tag:{} alias={}".format( alias["tag"], alias), "alias_replace")
            tagstart = pStr.find(alias["tag"])
            if tagstart >= 0:
                if alias["declid"] is not None:
                    decltag = s.aliasList[alias["declid"]]

                    tagstr = decltag["obj"].text
                    e_s = tagstr.find(s.TAG_S_EVAL)
                    if e_s >= 0:
                        s.eval_replace(decltag["obj"], decltag["objtype"])
                        tagstr = decltag["obj"].text
                    #
                    i_s = tagstr.find(s.TAG_S_ALIASDECL, 0)
                    i_e = tagstr.find(s.TAG_E_ALIASDECL, i_s)
                    # sqlmsg("alias_replace-ss i_s: {} , tagstr:{} declid:{}".format(i_s, tagstr, alias["declid"]), "alias_replacess")
                    if i_s >= 0:
                        tagbody = tagstr[i_s + len(s.TAG_S_ALIASDECL) : i_e]
                        parts = tagbody.split(s.TAG_D_ALIASDECL, 1)
                        ### These are replacement operators
                        for op in alias["ops"]:
                            if op == "lower":
                                parts[1] = str(parts[1]).lower()
                            #
                            if op == "upper":
                                parts[1] = str(parts[1]).upper()
                            #
                            if op in ("initcap", "capitalize"):
                                parts[1] = str(parts[1]).capitalize()
                            #
                        #
                        pStr = pStr.replace(alias["tag"], parts[1])
                    #
                    # sqlmsg("alias_replace- val:{} pStr:{}".format(alias, pStr), "alias_replace2")
                    # pStr = pStr.replace(alias["tag"],decltag["body"])
                #
            #
        #

        return pStr

    #

    def alias_clean(s):
        debug = []
        for alias in s.aliasList:

            if alias["obj"]:
                tagstr = alias["obj"].text
                a = alias.copy()
                a["obj"] = None
                debug.append(a)

                i_s = tagstr.find(s.TAG_S_ALIASDECL, 0)
                i_e = tagstr.find(s.TAG_D_ALIASDECL, i_s + len(s.TAG_S_ALIASDECL))
                i_f = tagstr.find(s.TAG_E_ALIASDECL, i_s)
                # sqlmsg("alias_replace- tagstr:{} i_s:{} i_e:{}".format(tagstr, i_s, i_e), "alias_replace2")
                if i_s >= 0 and i_e > i_s and i_e <= i_f:
                    tagstr = tagstr[i_e + len(s.TAG_D_ALIASDECL) :]
                #

                # tagstr = tagstr.replace(s.TAG_S_ALIASDECL,"")
                tagstr = tagstr.replace(s.TAG_E_ALIASDECL, "")
                if alias["type"] == "use":
                    tagstr = tagstr.replace(alias["tag"], "")
                #
                alias["obj"].text = tagstr
            #
        #

        # sqlmsg("alias_clean- debug:{} ".format(json.dumps(debug)), "alias_clean")

    #

    def str_value_replace(
        s,
        p_str,
        p_idx=None,
        p_fields=None,
        p_type="",
        p_tblid="",
        p_colidx=0,
        p_htmlnode=None,
    ):
        """Scans through string and replace from the given values."""
        if p_idx is None:
            p_idx = 0
        opdata = {
            "error": None,
            "rmblnk": 0,
            "mergedtags": [],
            "str": p_str,
            "rmblnkcol": 0,
        }

        if len(p_str) < 3:
            opdata["error"] = "String has no data."
            return opdata
        #

        if not p_str.find(s.TAG_S) >= 0:  # abort this line/object, we have not tags.
            opdata["error"] = "No tags found."
            return opdata
        #
        fields = p_fields
        if fields is None:
            fields = s.inputdata.get("fields")
        #

        if fields is None or type(fields) is not dict:
            # plpy.warning("No fields to replace in tag_value_replace. {} {}".format(fields,tags))
            opdata["error"] = "No fields given."
            return opdata
        #

        lasttag = None
        for tag in fields:
            if type(fields.get(tag)) is not dict:
                raise MergeToolError(
                    "Expecting tag field to have a type and value property.",
                    "tag_value_replace",
                )
            #
            found = False
            lasttag = tag
            val = None
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
                    plpy.warning(
                        "Field replace Exception: {} Detail: i:{} tag:{},datalist:{}".format(
                            e, p_idx, tagtype, tagvalue
                        )
                    )
                    pass
                #
            else:
                val = tagvalue
                if tagvalue != "":
                    found = True
                #
            #
            if tagtype == s.TTYP_PIC:
                if s.HTML:

                    def findtagdetails(p_wtype):
                        nonlocal p_str
                        val = ""
                        i_ts = p_str.find(tag)
                        i_sts = p_str.find("[^", i_ts)
                        # sqlmsg("add [2] findtagdetails for tag:{} , i_ts:{}, i_sts:{},val:{}".format(tag,i_ts,i_sts,p_str[i_ts:i_sts]), p_title="pl_mailmerge", p_type="local notice")
                        if i_sts >= 0 and i_sts >= i_ts:
                            i_ste = p_str.find("^]", i_sts) + 2

                            i_ws = p_str.find(p_wtype.lower() + "=", i_sts, i_ste)
                            if i_ws < 0:
                                i_ws = p_str.find(p_wtype.upper() + "=", i_sts, i_ste)
                            #
                            i_we = p_str.find(",", i_ws, i_ste)
                            if i_we < 0:
                                i_we = p_str.find("^", i_ws, i_ste)
                            #

                            if i_we >= 0 and i_ws >= 0:
                                val = p_str[i_ws + 2 : i_we]
                            #

                            p_str = p_str[:i_sts] + p_str[i_ste:]

                            # sqlmsg("add [i] findtagdetails for tag:{} , i_sts:{}, i_ste:{},i_ws:{},i_we:{}  ,val:{},run:{}".format(tag,i_sts,i_ste,i_ws,i_we,val,p_str), p_title="pl_mailmerge", p_type="local notice")
                        #
                        if val.isalnum():
                            return float(val)
                        else:
                            return 0
                        #

                    #

                    h = float(fields.get(tag).get("h", "0.5")) * 96
                    w = float(fields.get(tag).get("w", "2")) * 96

                    # sqlmsg("test findtagdetails for tag:{} , i_ts:{} , i_sts:{} ,val:{}".format(tag,i_ts,i_sts,p_str), p_title="pl_mailmerge", p_type="local notice")
                    user_w = findtagdetails("w")
                    user_h = findtagdetails("h")
                    if user_w > 0 or user_h > 0:
                        w = user_w
                        h = user_h
                    #

                    # width="{}%" height:"{}%"
                    style = ""
                    if w > 0:
                        style += "width: {}px;".format(w)
                    if h > 0:
                        style += "height: {}px;".format(h)
                    # sqlmsg("style for tag:{} ,style: {}, val:{}".format(tag,style,p_str), p_title="pl_mailmerge", p_type="local notice")

                    imgstr = '<img class="mergedimage" src="data:image/png;base64,{}" style="{}"  />'.format(
                        val, style
                    )

                    if p_htmlnode is not None:
                        imgnode = LET.XML(imgstr)
                        p_htmlnode.append(imgnode)
                        p_str = p_str.replace(tag, "")
                        continue
                    else:
                        p_str = p_str.replace(tag, imgstr)
                    #

                #
                continue
            #

            ### E-Sign tags
            if tag.lower().find(s.STR_TAG_ESIGN) > -1:
                if s.HTML and tag.find("esign_url") >= 0:
                    val = '<a href="{0}">{0}</a>'.format(val)
                #
            #

            if tagtype == s.TTYP_TBLROOT:
                continue  # we do not handle inner values here.
            if tagtype == s.TTYP_CONDITIONAL:
                if p_str.find(tag) >= 0:
                    if tryInt(tagvalue) > 0:
                        if tag.find("rmr!") >= 0:
                            val = s.TAG_COND_RMR
                            # sqlmsg("Will Remove Row: {} , {}, p_str={}".format(tag, tagvalue, p_str))
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
            #
            if type(val) is not str:
                plpy.warning("Expected string type value. {}".format(type(val), val))
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
                if p_str.find(tag) >= 0:
                    opdata["mergedtags"].append(tag)
                #
            #

            p_str = p_str.replace(tag, val)

            # plpy.notice("Tag replace: {}={}  obj:{}".format(tag, val, p_obj.text))
            #
        #

        opdata["str"] = p_str
        return opdata

    #

    def replace_str(s, p_str, p_values, p_idx=None):
        """Scans through a string for merge tags and replaces tags with values."""
        # if this is slow, we need to find a better alogrithm

        if p_idx is None:
            p_idx = 0
        for t in p_values:
            if type(p_values.get(t, "")) is dict:
                for k in p_values.get(t):
                    try:
                        if type(p_values.get(t).get(k)) is list:
                            val = p_values.get(t).get(k)[p_idx]
                            if val is None:
                                val = ""
                            p_str = istrreplace(p_str, k, val)  # p_str.replace(k, val)
                        else:
                            p_str = istrreplace(
                                p_str, k, p_values.get(t).get(k)
                            )  # p_str.replace(k, p_values.get(t).get(k))
                        #
                    except Exception as e:
                        plpy.notice(
                            "Exception: {} Detail: i:{} t:{},k{},v:{}".format(
                                e, p_idx, t, k, p_values.get(t).get(k)
                            )
                        )
                        pass
                    #
                #
            else:
                plpy.notice(
                    "Replacing: {} - {} -> {}".format(p_str, t, p_values.get(t))
                )
                p_str = istrreplace(
                    p_str, t, p_values.get(t)
                )  # p_str.replace(t, p_values.get(t))
            #
        #
        return p_str

    #

    @staticmethod
    def safe_eval(expr, variables):

        try:

            _safe_names = {"None": None, "True": True, "False": False}
            _safe_nodes = [
                "Add",
                "And",
                "BinOp",
                "BitAnd",
                "BitOr",
                "BitXor",
                "BoolOp",
                "Compare",
                "Dict",
                "Eq",
                "Expr",
                "Expression",
                "For",
                "Gt",
                "GtE",
                "Is",
                "In",
                "IsNot",
                "LShift",
                "List",
                "Load",
                "Lt",
                "LtE",
                "Mod",
                "Name",
                "Not",
                "NotEq",
                "NotIn",
                "Num",
                "Or",
                "RShift",
                "Set",
                "Slice",
                "Str",
                "Sub",
                "Tuple",
                "UAdd",
                "USub",
                "UnaryOp",
                "boolop",
                "cmpop",
                "expr",
                "expr_context",
                "operator",
                "slice",
                "unaryop",
                "print",
                ",",
                "*",
                "+",
                "-",
                "Mult",
                "Div",
                "Add",
                "repeat",
                "BT",
                "find",
                "contains",
                "Attribute",
                "str",
                "sub",
                "Call",
                "!=",
                "==",
                "Constant",
                "if",
                "else",
                ":",
                "int",
                "float",
                "str",
                "n2w",
                "nint",
                "ndec",
                "IfExp",
                "ElseExp",
            ]
            node = ast.parse(expr, mode="eval")
            for subnode in ast.walk(node):
                subnode_name = type(subnode).__name__
                # sqlmsg("Nodes breakdown: subnode_name={}, type:{}".format(subnode_name, type(subnode)), p_type="local notice")
                if isinstance(subnode, ast.Name):
                    # sqlmsg("Node data: subnode_name={}, subnodeid={} expr={}".format(subnode_name, subnode.id, expr), p_type="local notice")
                    if (
                        subnode.id is not None
                        and subnode.id not in _safe_names
                        and subnode.id not in _safe_nodes
                        and subnode.id not in variables
                    ):
                        raise ValueError(
                            "Unsafe expression {}. contains {}".format(expr, subnode.id)
                        )
                    #
                if subnode_name not in _safe_nodes:
                    # sqlmsg("Node data: subnode_name={} expr={}".format(subnode_name, expr), p_type="local notice")
                    raise ValueError(
                        "Unsafe expression {}. contains {}".format(expr, subnode_name)
                    )
                #
            #
            tools = {}

            def contains(p_str, p_val):
                # sqlmsg("Contains --- p_tuple:{} p_val:{} expr:{} ".format(p_str, p_val,expr))
                if str(p_str).find(str(p_val)) >= 0:
                    return True
                #
                return False

            #
            variables["contains"] = contains
            variables["n2w"] = n2w
            variables["nint"] = nint
            variables["ndec"] = ndec

            result = eval(expr, {"__builtins__": __builtins__}, variables)
        except Exception as e:
            sqlmsg(
                "safe_eval {} expr:{} vars:{}".format(e, expr, variables),
                p_type="local notice",
            )
            result = "Exp Error: {}  expr:{} vars:{}".format(e, expr, variables)
            raise Exception("Exp Error: {}  expr:{} vars:{}".format(e, expr, variables))
        #
        return result

    #

    def evelstr_replace_money(s, p_str):
        newstr = p_str
        try:
            if newstr.find("R") >= 0:
                newstr = newstr.replace("\xa0", "")
            #
            for i in range(0, 10):
                rstart = newstr.find("R")
                # sqlmsg("++ evelstr_replace_money: p_str={} result={}  i={}".format(rstart, newstr[rstart+1] , i ), p_type="local notice")

                if rstart >= 0 and (
                    isFloat(newstr[rstart + 1]) or isFloat(newstr[rstart + 1])
                ):
                    newstr = newstr[:rstart] + newstr[rstart + 1 :]
                    cstart = newstr.find(",")
                    if cstart > rstart and (
                        isFloat(newstr[rstart + 1])
                        or isFloat(newstr[rstart + 2])
                        or isFloat(newstr[rstart - 1])
                    ):
                        # sqlmsg("++ Comma Replace: {} newstr={}  {},{}".format(i, newstr, newstr[:cstart], newstr[cstart+1:]), p_type="local notice")
                        newstr = newstr[:cstart] + newstr[cstart + 1 :]
                    #
                #
            #
        except Exception as e:
            sqlmsg("evelstr_replace_money {} ".format(e), p_type="local error")
            return p_str
        #
        return newstr

    #

    def evalstr(s, p_str, p_def, p_obj=None, p_rules={}):
        result = ""
        result_pre = ""
        p_str = p_str.replace(chr(8221), '"')
        try:
            result_pre = s.evelstr_replace_money(p_str)
            result = MergeTool.safe_eval(result_pre.lstrip(" "), {})
            # sqlmsg("++ Merge Eval: p_rules={} ".format(p_rules), p_type="local notice")

            try:
                if p_rules is None:
                    p_rules = {}
                #
                ### The purpose of this complex post replace is to make the xx.0 or xx.1 to xx.00 or xx.01
                if isinstance(result, str) and result.find(".") == len(result) - 2:
                    fmt = "{:03.2f}"
                    if p_rules["noformat"]:
                        fmt = "{:03.2f}"
                    #
                    result = fmt.format(float(result))

                elif isinstance(result, float):
                    fmt = "{:03.2f}"
                    if p_rules["noformat"]:
                        fmt = "{:03.2f}"
                    #
                    f_result = str(round(result, 2))

                    if f_result.find(".") == len(f_result) - 2:
                        result = fmt.format(float(f_result))
                    #
                    if len(f_result) - f_result.find(".") <= 3:
                        result = fmt.format(float(f_result))
                    #

                    # sqlmsg("__-__Merge Eval: p_str={} result={} f_result={} , {}, {}".format(p_str, result, f_result, len(f_result), f_result.find('.')), p_type="local notice")

                elif isFloat(result):
                    f_result = str(round(result, 2))
                    if f_result.find(".") == len(f_result) - 2:
                        result = "{:.}".format(float(f_result))
                    #
                #

            except Exception as e:
                sqlmsg(
                    "Merge Eval Post process Error: {} \r\nStr={}".format(e, p_str),
                    p_type="local notice",
                )
            #
            # sqlmsg("Merge Eval: {} \r\result_pre={} result={}".format(p_str, result_pre.lstrip(" "),result ), p_type="local notice")

        except Exception as e:
            sqlmsg(
                "Merge Eval Error: {} \r\nStr={} Result_pre={} Result={}".format(
                    e, p_str, result_pre, result
                ),
                p_type="local notice",
            )
            if p_obj is not None:
                s.post_format.append({"tag": "eval", "obj": p_obj, "type": "error"})
            #
            raise Exception(
                "Eval Error: in {}: {} Result_pre={} Result={}".format(p_str, e)
            )
        #
        return str(result)

    #

    def eval_replace(s, p_obj, p_type="", p_tableid=""):
        eval_tag_start = 0
        eval_tag_end = 0
        found = False
        canBlank = True
        err = False
        if isinstance(p_obj, docx.text.run.Run):
            m_str = p_obj.text
            p_type = "docx"
        elif hasattr(p_obj, "text"):
            m_str = p_obj.text
            p_type = "docx"
        else:
            m_str = p_obj
        #
        # if m_str.find(s.TAG_S_EVAL) >= 0:
        #  sqlmsg("Eval[tag] tag:{}".format(m_str))
        #
        if p_type in ("html", "txt"):
            m_str = html_replace(m_str)
        #

        evalAliasBegin = m_str.find(s.TAG_S_ALIASDECL)
        if m_str.find(s.TAG_S_ALIASUSE) >= 0:
            bk = m_str
            m_str = s.alias_replace(m_str)
            # sqlmsg("Eval alias - pre:{} post:{}".format(bk,m_str))
        #

        for i in range(50):
            eval_tag_start = m_str.find(s.TAG_S_EVAL, 0)
            eval_tag_end = m_str.find(s.TAG_E_EVAL, eval_tag_start)

            inAlias = evalAliasBegin >= 0 and eval_tag_start >= 0
            cleanresult = False
            exitloop = 0
            noFormat = inAlias
            try:
                # sqlmsg("Eval[post] S:{} E:{} Str:{}".format(eval_tag_start, eval_tag_end, m_str[eval_tag_start+len(s.TAG_S_EVAL):eval_tag_end]))
                if eval_tag_start >= 0 and eval_tag_end >= eval_tag_start:
                    exprstr = m_str[eval_tag_start + len(s.TAG_S_EVAL) : eval_tag_end]
                    posttag_end = exprstr.find(s.TAG_S_POSTOP)
                    if exprstr.find(s.TAG_E_POSTOP) > posttag_end and posttag_end > 0:
                        continue
                    #

                    opstrings = []
                    i_op_start = m_str.find("#", eval_tag_start, eval_tag_end)
                    if i_op_start > 0:
                        opend = m_str.find("#", i_op_start + 1)
                        # opstr = m_str[i_op_start+1:opend]
                        opstrings = m_str[i_op_start + 1 : eval_tag_end].split("#")
                        exprstr = m_str[eval_tag_start + len(s.TAG_S_EVAL) : i_op_start]
                        for t in opstrings:
                            if t.find("fmt") >= 0:
                                noFormat = True
                            #
                        #
                        # sqlmsg("Eval[Operator] exprstr:[{}] opstr:{} [{}] Str:[{}]".format(exprstr, i_op_start, opstr, m_str[eval_tag_start+len(s.TAG_S_EVAL):eval_tag_end] ))
                    #
                    # sqlmsg("Eval[str] opstrings={} exprstr={}".format(opstrings, exprstr))
                    newstr = s.evalstr(exprstr, "", p_obj, {"noformat": noFormat})
                    exprtrue = newstr.strip(" ").lower() in ("1", "true")
                    exprfalse = not exprtrue

                    # sqlmsg("Eval[str] opstrings:[{}] exprtrue:{} exprstr:[{}]".format(opstrings,exprtrue, exprstr ))
                    def getOpStr(pKey):
                        for opstr in opstrings:
                            strstartpos = opstr.lower().find(pKey)
                            if strstartpos >= 0:
                                istart = strstartpos + len(pKey)
                                newstr = opstr[istart:]
                                return newstr
                            #
                        #
                        return ""

                    #

                    for opstr in opstrings:
                        strstartpos = opstr.lower().find("str:")

                        if strstartpos >= 0:
                            if exprtrue:
                                istart = strstartpos + len("str:")
                                newstr = opstr[istart:]

                                m_str = (
                                    m_str[:eval_tag_start]
                                    + newstr
                                    + m_str[eval_tag_end + len(s.TAG_E_EVAL) :]
                                )
                                found = True
                            else:
                                newstr = ""
                                if exprfalse:
                                    newstr = getOpStr("str!:")
                                #
                                m_str = (
                                    m_str[:eval_tag_start]
                                    + newstr
                                    + m_str[eval_tag_end + len(s.TAG_E_EVAL) :]
                                )
                                found = True
                            #

                            # sqlmsg("Eval[str2] newstr:[{}] exprtrue:{} exprstr:[{}] m_str:[{}]".format(newstr,exprtrue, exprstr,m_str ))

                            exitloop = 1

                        elif opstr.find("rmp") >= 0:
                            canBlank = False
                            cleanresult = True

                            if isinstance(p_obj, docx.text.run.Run) and (
                                exprtrue or exprfalse and opstr.find("rmp!")
                            ):
                                par = p_obj._parent
                                gran = par._parent
                                par.text = ""
                                par.clear()
                            #
                        elif (
                            opstr.find("rmt") >= 0
                            or opstr.find("rmc") >= 0
                            or opstr.find("rmr") >= 0
                            or opstr.find("rmo") >= 0
                        ):
                            inverted = (
                                opstr.find("rmt!") >= 0
                                or opstr.find("rmc!") >= 0
                                or opstr.find("rmr!") >= 0
                                or opstr.find("rmo!") >= 0
                            )
                            cleanresult = True
                            canBlank = False
                            tbl = None
                            if isinstance(p_obj, docx.text.run.Run) and (
                                exprtrue or inverted and exprfalse
                            ):
                                par = p_obj._parent
                                gran = par._parent
                                if isinstance(par, docx.table._Cell):
                                    tbl = par._parent
                                elif isinstance(gran, docx.table._Cell):
                                    tbl = gran._parent
                                #
                                if tbl is not None:
                                    if opstr.find("rmc") >= 0:
                                        gran.text = ""
                                    #
                                    if opstr.find("rmr") >= 0:
                                        rowcnt = 0
                                        totalrows = len(tbl.rows)
                                        for r in tbl.rows:
                                            rowcnt += 1
                                            for c in r.cells:
                                                if id(c._element) == id(gran._element):
                                                    parent = r._tr.find("..")
                                                    s.table_removerow_list.append(
                                                        {
                                                            "xmlptr": tbl,
                                                            "rowptr": r,
                                                            "row": rowcnt,
                                                            "total": totalrows,
                                                            "tableid": p_tableid,
                                                        }
                                                    )

                                                    if parent is not None:
                                                        parent.remove(r._tr)
                                                        break
                                                    #
                                                #
                                            #
                                        #
                                    #
                                    if opstr.find("rmo") >= 0:
                                        canBlank = False
                                        for col in tbl.columns:
                                            found = False
                                            for c in col.cells:
                                                if id(c._element) == id(gran._element):
                                                    found = True
                                                    break
                                                #
                                            #
                                            if found:
                                                for c in col.cells:
                                                    c.text = ""
                                                #
                                            #
                                        #
                                    #
                                    if opstr.find("rmt") >= 0:
                                        p = tbl._parent
                                        if p is not None:
                                            p._element.remove(tbl._element)
                                        #
                                    #
                                #
                            #
                        elif opstr.find("border") >= 0:
                            canBlank = False
                            inverted = opstr.find("border!") >= 0
                            borderopts = opstr[len("border") + 1 :]
                            cleanresult = True
                            if isinstance(p_obj, docx.text.run.Run) and (
                                exprtrue or inverted and exprfalse
                            ):
                                par = p_obj._parent
                                gran = par._parent
                                if isinstance(gran, docx.table._Cell):
                                    bde = ""
                                    try:
                                        xcell = gran._element

                                        fmt = """<?xml version="1.0"?>
                    <w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                    """

                                        if borderopts.find("db") >= 0:
                                            fmt += """<w:bottom w:val="double" w:color="auto" w:space="0" w:sz="4"/>"""
                                        #
                                        elif borderopts.find("b") >= 0:
                                            fmt += """<w:bottom w:val="single" w:color="auto" w:space="0" w:sz="4"/>"""
                                        #

                                        if borderopts.find("dt") >= 0:
                                            fmt += """<w:top w:val="double" w:color="auto" w:space="0" w:sz="4"/>"""
                                        #
                                        elif borderopts.find("t") >= 0:
                                            fmt += """<w:top w:val="single" w:color="auto" w:space="0" w:sz="4"/>"""
                                        #

                                        if borderopts.find("dl") >= 0:
                                            fmt += """<w:left w:val="sindoublegle" w:color="auto" w:space="0" w:sz="4"/>"""
                                        #
                                        elif borderopts.find("l") >= 0:
                                            fmt += """<w:left w:val="single" w:color="auto" w:space="0" w:sz="4"/>"""
                                        #

                                        if borderopts.find("dr") >= 0:
                                            fmt += """<w:right w:val="double" w:color="auto" w:space="0" w:sz="4"/>"""
                                        #
                                        elif borderopts.find("r") >= 0:
                                            fmt += """<w:right w:val="single" w:color="auto" w:space="0" w:sz="4"/>"""
                                        #

                                        fmt += """</w:tcBorders>"""
                                        bde = LET.fromstring(fmt)
                                        ptr = xcell.find("tcPr")
                                        if ptr is None:
                                            ptr = LET.fromstring(
                                                """<?xml version="1.0"?>
                    <w:tcPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                    </w:tcPr>
                    """
                                            )
                                            xcell.append(ptr)
                                        #
                                        ptr.append(bde)
                                    except Exception as e:
                                        raise Exception(
                                            "Failed to change border settings. {}\r\n{}".format(
                                                e, bde
                                            )
                                        )
                                    #
                                #
                            #
                        elif opstr.find("addrows") >= 0 and exprtrue:
                            canBlank = False
                            borderopts = opstr[len("addrows") + 1 :]
                            cleanresult = True

                            newrowcnt = int(borderopts.strip(" "))
                            if isinstance(p_obj, docx.text.run.Run):
                                par = p_obj._parent
                                gran = par._parent
                                if isinstance(gran, docx.table._Cell):
                                    tbl = gran._parent
                                    for i in range(newrowcnt):
                                        tbl.add_row()
                                    #
                                #
                            #
                        elif opstr.find("bold") >= 0:
                            canBlank = False
                            inverted = opstr.find("bold!") >= 0
                            borderopts = opstr[len("bold") + 1 :]
                            cleanresult = True
                            if isinstance(p_obj, docx.text.run.Run) and (
                                exprtrue or inverted and exprfalse
                            ):
                                par = p_obj._parent
                                font = p_obj.font
                                if font is not None:
                                    font.bold = True
                                #
                            #
                        elif opstr.find("italic") >= 0:
                            canBlank = False
                            inverted = opstr.find("italic!") >= 0
                            borderopts = opstr[len("italic") + 1 :]
                            cleanresult = True
                            if isinstance(p_obj, docx.text.run.Run) and (
                                exprtrue or inverted and exprfalse
                            ):
                                par = p_obj._parent
                                font = p_obj.font
                                if font is not None:
                                    font.italic = True
                                #
                            #
                        elif opstr.find("underline") >= 0:
                            canBlank = False
                            inverted = opstr.find("underline!") >= 0
                            borderopts = opstr[len("underline") + 1 :]
                            cleanresult = True
                            if isinstance(p_obj, docx.text.run.Run) and (
                                exprtrue or inverted and exprfalse
                            ):
                                par = p_obj._parent
                                font = p_obj.font
                                if font is not None:
                                    font.underline = True
                                #
                            #
                        #
                        # -----------------------------------------------------------[CHANGE ID: 29879]  25/10/2021 16:03 Begin
                        elif opstr.find("hide") >= 0:
                            inverted = opstr.find("hide!") >= 0
                            opts = opstr[len("hide") + 1 :]
                            times = tryInt(opts)
                            cleanresult = True
                            # sqlmsg("Hide {} for {} op: {}".format(m_str,times, opts ))
                            if isinstance(p_obj, docx.text.run.Run):
                                par = p_obj._parent
                                sib = None
                                for i in range(times + 1):
                                    if inverted:
                                        sib = p_obj._element.getnext()
                                    else:
                                        sib = p_obj._element.getprevious()
                                    #
                                    if sib is None:
                                        break
                                    #
                                    par.remove(sib)
                                #
                                p_obj.text = ""
                            #
                        #
                        # -----------------------------------------------------------[CHANGE ID: 29879]  25/10/2021 16:03 End
                        elif opstr.find("fmt") >= 0:
                            borderopts = opstr[len("fmt") + 1 :]
                            ops = borderopts.split(",")
                            cleanresult = True

                            if isinstance(p_obj, docx.text.run.Run):
                                tmpstr = newstr
                                if len(ops) > 1:
                                    fmt = "{:" + ops[0] + ",." + ops[1] + "f}"
                                elif len(ops) == 1:
                                    fmt = "{:" + ops[0] + ",.2f}"
                                else:
                                    fmt = "{:.2f}"
                                #
                                try:

                                    floatval = tryFloat(tmpstr, "0")
                                    if floatval > 0:
                                        tmpstr = fmt.format(floatval)
                                        sqlmsg(
                                            "Eval format floatval:{} opstrings:{} borderopts:{} tmpstr: {}".format(
                                                floatval, opstrings, borderopts, tmpstr
                                            ),
                                            p_type="local notice",
                                        )

                                        newstr = tmpstr
                                        m_str = tmpstr
                                        found = True
                                        break
                                    #
                                except Exception as e:
                                    raise "Failed to format string to float: {}".format(
                                        p_obj.text
                                    )
                                #
                            #
                    # end for ops

                    if exitloop == 1:
                        continue
                    if exitloop == 2:
                        break

                    if cleanresult:
                        newstr = ""
                    #
                    m_str = (
                        m_str[:eval_tag_start]
                        + newstr
                        + m_str[eval_tag_end + len(s.TAG_E_EVAL) :]
                    )

                    found = True
                else:
                    if eval_tag_start >= 0:
                        sqlmsg(
                            "Eval open tag {} found without ending tag in this run->{}".format(
                                s.TAG_S_EVAL, m_str
                            ),
                            p_type="local notice",
                        )
                    #
                    break
                #
            except Exception as e:
                if glo_showerrors:
                    m_str = '[! Expression Failed: {} in "{}"!]'.format(e, exprstr)
                else:
                    m_str = ""
                #
                err = True
            #
        #
        if found and isinstance(p_obj, docx.text.run.Run):
            # sqlmsg("Blank Text: {}".format(m_str ))
            # if canBlank and len(m_str) < 2:
            #   m_str = "."
            #   s.post_format.append({"tag": 'eval', "obj": p_obj, "type": 'default'})
            # #
            if len(m_str) < 2:
                m_str = ""
            #
            p_obj.text = m_str
        elif err and isinstance(p_obj, docx.text.run.Run):
            p_obj.text = m_str
        #
        # sqlmsg("Eval[Return] m_str:[{}]".format(m_str ))
        return m_str

    #


#
