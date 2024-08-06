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
from ..tags import (
    TAG_S_SP,
    TAG_E_SP,
    TAG_S,
    TAG_E,
    TAG_S_OLD,
    TAG_E_OLD,
    TAG_S_EVAL,
    TAG_E_EVAL,
    TAG_S_INPLACE,
    TAG_E_INPLACE,
    TAG_S_ALIASDECL,
    TAG_E_ALIASDECL,
    TAG_D_ALIASDECL,
    TAG_S_ALIASUSE,
    TAG_E_ALIASUSE,
    TAG_S_POSTOP,
    TAG_E_POSTOP,
    TAG_S_TBLM,
    TAG_E_TBLM,
    TAG_ERR,
    STR_ERR,
    TAG_SPLT,
    TAG_RULESPLT,
    STR_TAG_RMBD,
    STR_TAG_RMBC,
    STR_TAG_RMDR,
    STR_TAG_FREETBL,
    STR_TAG_ESIGN,
    STR_TAG_LIMITROW,
    STR_TAG_TBLSEARCH_RM,
    TTYP_ROOT,
    TTYP_FIELD,
    TTYP_TBLFIELD,
    TTYP_TBLROOT,
    TTYP_AGGFIELD,
    TTYP_PIC,
    TTYP_SPECIAL,
    TTYP_FILTER,
    TTYP_CONDITIONAL,
    TTYP_DOCREPLACE,
    TTYP_EMBEDHTML,
    OP_TAGS,
    CELL_OP_TAGS,
    TAG_COND_RMR,
    TAG_COND_RMC,
    TAG_COND_CLR,
    TAG_COND_RTBL,
    TAG_COND_TBL_DISTINCT,
    TAG_COND_RMP,
)
import ast

if "plpy" not in globals():
    import pgsql_utils.pgproc.wrap_plpy as plpy
#

class MergeTool(object):
    """Bitech Merge Tool"""

    DEBUG = False

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
        s.show_errors = True

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

        search_str(p_str, TAG_S, TAG_E)
        if p_inc_old:
            search_str(p_str, TAG_S_OLD, TAG_E_OLD)
            search_str(p_str, TAG_S_SP, TAG_E_SP)
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
                    e_s = tagstr.find(TAG_S_EVAL)
                    if e_s >= 0:
                        s.eval_replace(decltag["obj"], decltag["objtype"])
                        tagstr = decltag["obj"].text
                    #
                    i_s = tagstr.find(TAG_S_ALIASDECL, 0)
                    i_e = tagstr.find(TAG_E_ALIASDECL, i_s)
                    # sqlmsg("alias_replace-ss i_s: {} , tagstr:{} declid:{}".format(i_s, tagstr, alias["declid"]), "alias_replacess")
                    if i_s >= 0:
                        tagbody = tagstr[i_s + len(TAG_S_ALIASDECL) : i_e]
                        parts = tagbody.split(TAG_D_ALIASDECL, 1)
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

                i_s = tagstr.find(TAG_S_ALIASDECL, 0)
                i_e = tagstr.find(TAG_D_ALIASDECL, i_s + len(TAG_S_ALIASDECL))
                i_f = tagstr.find(TAG_E_ALIASDECL, i_s)
                # sqlmsg("alias_replace- tagstr:{} i_s:{} i_e:{}".format(tagstr, i_s, i_e), "alias_replace2")
                if i_s >= 0 and i_e > i_s and i_e <= i_f:
                    tagstr = tagstr[i_e + len(TAG_D_ALIASDECL) :]
                #

                # tagstr = tagstr.replace(s.TAG_S_ALIASDECL,"")
                tagstr = tagstr.replace(TAG_E_ALIASDECL, "")
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

        if not p_str.find(TAG_S) >= 0:  # abort this line/object, we have not tags.
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
            if tagtype == TTYP_PIC:
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
            if tag.lower().find(STR_TAG_ESIGN) > -1:
                if s.HTML and tag.find("esign_url") >= 0:
                    val = '<a href="{0}">{0}</a>'.format(val)
                #
            #

            if tagtype == TTYP_TBLROOT:
                continue  # we do not handle inner values here.
            if tagtype == TTYP_CONDITIONAL:
                if p_str.find(tag) >= 0:
                    if tryInt(tagvalue) > 0:
                        if tag.find("rmr!") >= 0:
                            val = TAG_COND_RMR
                            # sqlmsg("Will Remove Row: {} , {}, p_str={}".format(tag, tagvalue, p_str))
                        elif tag.find("rmc!") >= 0:
                            val = TAG_COND_RMC
                        elif tag.find("cc!") >= 0:
                            val = TAG_COND_CLR
                        elif tag.find("!rmp!") >= 0:
                            val = TAG_COND_RMP
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


#
