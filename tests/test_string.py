import unittest

import pgsql_utils.util.string as imp_string


class TestStringOperations(unittest.TestCase):
    def test_n2w(self):
        self.assertEqual(imp_string.n2w(6), "Six")

    def test_ordinalstr(self):
        self.assertEqual(imp_string.ordinalstr(6), "6th")


if __name__ == "__main__":
    unittest.main(verbosity=2)
