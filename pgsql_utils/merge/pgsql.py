### Copyright (c) 2024 Bitech Systems. All rights reserved.
### The code and materials in this repository are the exclusive property of Bitech Systems and its associated companies and are protected by copyright law.
### Please refer to the license details in the package.

from .exception import MergeToolWarning, MergeToolError
from .mdocx.merge import MergeToolDocX
from .html.merge import MergeToolHTML
from .text.merge import MergeToolTEXT
from ..pgproc.debug import sqlmsg, get_err_msg
import sys, os, traceback, json

if "plpy" not in globals():
    import pgsql_utils.pgproc.wrap_plpy as plpy
#


MODES = (
    "peek_docx",
    "peekall_docx",
    "peek_html",
    "peek_html",
    "peek_txt",
    "merge_docx",
    "merge_html",
    "merge_txt",
    "convert_docx",
    "convert_html",
    "convert_txt",
    "convertonly_docx",
    "fixanchor_docx",
    "mocktest_docx",
    "mocktest_html",
    "mocktest_txt",
    "mark_docx",
    "mark_html",
    "mark_txt",
)


def check_merge_inputs(
    obj, p_template_filename, p_template_blob, p_src_filename, p_jsonset=None
):
    """Check if the input data combination is correct"""

    try:
        if p_template_blob is not None and len(p_template_blob) > 1:
            obj.template_blob = p_template_blob
        #

        if p_template_filename is not None and len(p_template_filename) > 0:
            obj.srcfilename = p_template_filename
            p_template_filename = os.path.normpath(p_template_filename)

            if p_template_filename is None or not os.path.exists(p_template_filename):
                if obj.mode.lower().find("mocktest") < 0:
                    raise MergeToolError(
                        "The template file given does not exist. Filename: {}".format(
                            p_template_filename
                        )
                    )
                #
            #
        #

        if "merge" in obj.mode:
            if p_src_filename is not None and len(p_src_filename) > 0:
                p_src_filename = os.path.normpath(p_src_filename)
            #
        #
        if p_src_filename is not None and len(p_src_filename) > 0:
            obj.destfilename = p_src_filename
        #
        if p_jsonset is not None:
            obj.inputdata = json.loads(p_jsonset)
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        raise MergeToolError(
            "Invalid Input Data: {}".format(e), "check_merge_inputs", exc_tb
        )
    #


#


def execute_merge_sql(
    p_mode,
    p_template_filename,
    p_src_filename,
    p_jsonset,
    p_template_blob,
    p_programmode=0,
):
    """Execute the SQL statement"""
    p_retval = 0
    p_errmsg = ""
    p_result = ""
    p_file = None
    try:
        objMerge = None
        p_mode = p_mode.replace("text", "txt")
        if "docx" in p_mode.lower():
            objMerge = MergeToolDocX()
        elif "html" in p_mode.lower():
            objMerge = MergeToolHTML()
        elif "txt" in p_mode.lower():
            objMerge = MergeToolTEXT()
        else:
            raise MergeToolError("Invalid mode: {}".format(p_mode))

        if objMerge is not None:
            objMerge.mode = p_mode
            objMerge.programmode = p_programmode
            check_merge_inputs(
                objMerge, p_template_filename, p_template_blob, p_src_filename
            )
            p_result = objMerge.execute()
        #

    except MergeToolWarning as e:
        p_retval = 2
        p_errmsg = str(e)
        plpy.warning(e, traceback.format_tb(exc_tb, 30))
    except MergeToolError as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        p_retval = 1
        plpy.warning(e, traceback.format_tb(exc_tb, 30))
        p_errmsg = get_err_msg(
            p_errmsg=str(e),
            p_errdetail=str(traceback.format_tb(exc_tb, 30)),
            p_context=", at line " + str(exc_tb.tb_lineno),
        )

    except plpy.SPIError as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        p_retval = 1
        plpy.warning(e, traceback.format_tb(exc_tb, 30))
        p_errmsg = get_err_msg(
            p_errmsg=str(e),
            p_errdetail=str(traceback.format_tb(exc_tb, 30)),
            p_context=", at line " + str(exc_tb.tb_lineno),
        )

    except Exception as inst:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        p_errmsg = get_err_msg(
            p_errmsg=str(inst),
            p_errdetail=str(traceback.format_tb(exc_tb, 30)),
            p_context=", at line " + str(exc_tb.tb_lineno),
        )
        plpy.warning(inst, traceback.format_tb(exc_tb, 30))
        p_retval = 1
    except:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        p_errmsg = get_err_msg(
            p_errmsg=str(exc_type),
            p_errdetail=str(traceback.format_tb(exc_tb, 30)),
            p_context=", at line " + str(exc_tb.tb_lineno),
        )
        plpy.warning(traceback.format_tb(exc_tb, 30))
        p_retval = 1

    return [p_retval, p_errmsg, p_result, p_file]
