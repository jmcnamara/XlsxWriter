###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...utility import xl_col_to_name


class TestUtility(unittest.TestCase):
    """
    Test xl_col_to_name() utility function.

    """

    def test_xl_col_to_name(self):
        """Test xl_col_to_name()"""

        tests = [
            # col,  col string
            (0, 'A'),
            (1, 'B'),
            (2, 'C'),
            (9, 'J'),
            (24, 'Y'),
            (25, 'Z'),
            (26, 'AA'),
            (254, 'IU'),
            (255, 'IV'),
            (256, 'IW'),
            (16383, 'XFD'),
            (16384, 'XFE'),
        ]

        for col, string in tests:
            exp = string
            got = xl_col_to_name(col)
            self.assertEqual(got, exp)

    def test_xl_col_to_name_abs(self):
        """Test xl_col_to_name() with absolute references"""

        tests = [
            # col, col_abs, col string
            (0, 1, '$A'),
        ]

        for col, col_abs, string in tests:
            exp = string
            got = xl_col_to_name(col, col_abs)
            self.assertEqual(got, exp)
