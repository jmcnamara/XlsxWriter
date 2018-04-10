###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...utility import xl_range
from ...utility import xl_range_abs


class TestUtility(unittest.TestCase):
    """
    Test xl_range() utility function.

    """

    def test_xl_range(self):
        """Test xl_range()"""

        tests = [
            # first_row, first_col, last_row, last_col, Range
            (0, 0, 9, 0, 'A1:A10'),
            (1, 2, 8, 2, 'C2:C9'),
            (0, 0, 3, 4, 'A1:E4'),
            (0, 0, 0, 0, 'A1:A1'),
            (0, 0, 0, 1, 'A1:B1'),
            (0, 2, 0, 9, 'C1:J1'),
            (1, 0, 2, 0, 'A2:A3'),
            (9, 0, 1, 24, 'A10:Y2'),
            (7, 25, 9, 26, 'Z8:AA10'),
            (1, 254, 1, 255, 'IU2:IV2'),
            (1, 256, 0, 16383, 'IW2:XFD1'),
            (0, 0, 1048576, 16384, 'A1:XFE1048577'),
        ]

        for first_row, first_col, last_row, last_col, cell_range in tests:
            exp = cell_range
            got = xl_range(first_row, first_col, last_row, last_col)
            self.assertEqual(got, exp)

    def test_xl_range_abs(self):
        """Test xl_range_abs()"""

        tests = [
            # first_row, first_col, last_row, last_col, Range
            (0, 0, 9, 0, '$A$1:$A$10'),
            (1, 2, 8, 2, '$C$2:$C$9'),
            (0, 0, 3, 4, '$A$1:$E$4'),
            (0, 0, 0, 0, '$A$1:$A$1'),
            (0, 0, 0, 1, '$A$1:$B$1'),
            (0, 2, 0, 9, '$C$1:$J$1'),
            (1, 0, 2, 0, '$A$2:$A$3'),
            (9, 0, 1, 24, '$A$10:$Y$2'),
            (7, 25, 9, 26, '$Z$8:$AA$10'),
            (1, 254, 1, 255, '$IU$2:$IV$2'),
            (1, 256, 0, 16383, '$IW$2:$XFD$1'),
            (0, 0, 1048576, 16384, '$A$1:$XFE$1048577'),
        ]

        for first_row, first_col, last_row, last_col, cell_range in tests:
            exp = cell_range
            got = xl_range_abs(first_row, first_col, last_row, last_col)
            self.assertEqual(got, exp)
