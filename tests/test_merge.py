import unittest

import pgsql_utils.pgproc.fake_plpy as plpy
from pgsql_utils.merge.pgsql import execute_sql


class TestMergeOperations(unittest.TestCase):
    def test_proc_call(self):
        dat = execute_sql("peek_docx", "test", "test",None, None)
        self.assertTrue(True)



if __name__ == "__main__":
    unittest.main(verbosity=2)
