# Global vars

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
