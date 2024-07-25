import unittest

import pgsql_utils.pgproc.fake_plpy as plpy
from pgsql_utils.merge import execute_merge_sql


class TestMergeOperations(unittest.TestCase):
    def test_proc_call(self):
        dat = execute_merge_sql("peek_docx", "test", "test", None, None)
        self.assertTrue(True)


if __name__ == "__main__":
    unittest.main(verbosity=2)
