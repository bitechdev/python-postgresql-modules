import unittest
import os
import pgsql_utils.pgproc.wrap_plpy as plpy
from pgsql_utils.merge import execute_merge_sql
import json, io


class TestMergeOperations(unittest.TestCase):
    def test_proc_call_peek_file(self):
        filename = os.path.join(
            os.path.dirname(__file__), "data/Mail_Merge_Test_Template.docx"
        )
        resultfile = os.path.join(
            os.path.dirname(__file__), "data/Mail_Merge_Test_Template_peek_result.json"
        )
        print("Testing File: ", filename)

        [p_retval, p_errmsg, p_result, p_file] = execute_merge_sql(
            "peek_docx", filename, None, None, None
        )
        print("Results", p_retval, p_errmsg)

        if p_errmsg is not None:
            self.assertEqual(len(p_errmsg), 0, p_errmsg)

        self.assertEqual(p_retval, 0, p_errmsg)

        with open(resultfile, "r") as jf:
            tagsobj = json.load(jf)
            tagstr = json.dumps(tagsobj)
            self.assertEqual(p_result, tagstr, "Tags data does not match")
        #

    def test_proc_call_peek_blob(self):
        filename = os.path.join(
            os.path.dirname(__file__), "data/Mail_Merge_Test_Template.docx"
        )
        resultfile = os.path.join(
            os.path.dirname(__file__), "data/Mail_Merge_Test_Template_peek_result.json"
        )
        print("Testing File As Blob: ", filename)
        with open(filename, "rb") as blobfile:
            blobdata = blobfile.read()
            [p_retval, p_errmsg, p_result, p_file] = execute_merge_sql(
                "peek_docx", None, None, None, blobdata
            )
            print("Results", p_retval, p_errmsg)

            if p_errmsg is not None:
                self.assertEqual(len(p_errmsg), 0, p_errmsg)

            self.assertEqual(p_retval, 0, p_errmsg)

            with open(resultfile, "r") as jf:
                tagsobj = json.load(jf)
                tagstr = json.dumps(tagsobj)
                self.assertEqual(p_result, tagstr, "Tags data does not match")
            #
        #


if __name__ == "__main__":
    unittest.main(verbosity=2)
