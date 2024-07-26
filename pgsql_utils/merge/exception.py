import traceback


class MergeToolError(Exception):
    def __init__(s, p_msg, p_funcname="", p_traceback=None):
        if p_funcname != "":
            s.value = p_msg + " in " + p_funcname
        else:
            s.value = p_msg
        #
        if p_traceback is not None:
            s.value = "{} \r\nTraceback: \r\n{}".format(
                s.value, traceback.format_tb(p_traceback, 30)
            )
        #

    #
    def __str__(s):
        return s.value

    #


#


class MergeToolWarning(Exception):
    def __init__(s, p_msg):
        s.value = p_msg

    #
    def __str__(s):
        return s.value

    #


#
