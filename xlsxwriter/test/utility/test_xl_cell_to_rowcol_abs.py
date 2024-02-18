###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...utility import xl_cell_to_rowcol_abs


class TestUtility(unittest.TestCase):
    """
    Test xl_cell_to_rowcol_abs() utility function.

    """

    def test_xl_cell_to_rowcol_abs(self):
        """Test xl_cell_to_rowcol_abs()"""

        tests = [
            # row, col, A1 string
            (0, 0, "A1"),
            (0, 1, "B1"),
            (0, 2, "C1"),
            (0, 9, "J1"),
            (1, 0, "A2"),
            (2, 0, "A3"),
            (9, 0, "A10"),
            (1, 24, "Y2"),
            (7, 25, "Z8"),
            (9, 26, "AA10"),
            (1, 254, "IU2"),
            (1, 255, "IV2"),
            (1, 256, "IW2"),
            (0, 16383, "XFD1"),
            (1048576, 16384, "XFE1048577"),
        ]

        for row, col, string in tests:
            exp = (row, col, 0, 0)
            got = xl_cell_to_rowcol_abs(string)
            self.assertEqual(got, exp)

    def test_xl_cell_to_rowcol_abs_abs(self):
        """Test xl_cell_to_rowcol_abs() with absolute references"""

        tests = [
            # row, col, row_abs, col_abs, A1 string
            (0, 0, 0, 0, "A1"),
            (0, 0, 1, 0, "A$1"),
            (0, 0, 0, 1, "$A1"),
            (0, 0, 1, 1, "$A$1"),
        ]

        for row, col, row_abs, col_abs, string in tests:
            exp = (row, col, row_abs, col_abs)
            got = xl_cell_to_rowcol_abs(string)
            self.assertEqual(got, exp)
