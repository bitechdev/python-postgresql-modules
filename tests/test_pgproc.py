import unittest

import pgsql_utils.pgproc.fake_plpy as plpy
from pgsql_utils.pgproc.debug import get_err_msg, sqlmsg, set_plpy

set_plpy(plpy)

class TestPlpyOperations(unittest.TestCase):
    def test_notice(self):
        plpy.notice("Notice Test")
        self.assertTrue(True)

    def test_error(self):
        plpy.error("Error Test")
        self.assertTrue(True)

    def test_warning(self):
        plpy.warning("Warn Test")
        self.assertTrue(True)

    def test_get_err_msg(self):
        self.assertIsNotNone(get_err_msg("test", "msg", "detail", "context"))

    def test_sqlmsg(self):
        self.assertIsNone(sqlmsg("msg", "title", "local notice"))


if __name__ == "__main__":
    unittest.main(verbosity=2)
