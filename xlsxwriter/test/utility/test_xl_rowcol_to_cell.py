###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
import warnings
from ...utility import xl_rowcol_to_cell
from ...utility import xl_rowcol_to_cell_fast


class TestUtility(unittest.TestCase):
    """
    Test xl_rowcol_to_cell() utility function.

    """

    def test_xl_rowcol_to_cell(self):
        """Test xl_rowcol_to_cell()"""

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
            (0, -1, None),
            (-1, 0, None),
        ]

        for row, col, range in tests:
            exp = range
            got = xl_rowcol_to_cell(row, col)

            # Ignore the warnings for negative values.
            warnings.filterwarnings("ignore")

            self.assertEqual(got, exp)

    def test_xl_rowcol_to_cell_abs(self):
        """Test xl_rowcol_to_cell() with absolute references"""

        tests = [
            # row, col, row_abs, col_abs, A1 range
            (0, 0, True, False, "A$1"),
            (0, 0, False, True, "$A1"),
            (0, 0, True, True, "$A$1"),
            (-1, 0, 0, 0, None),
            (0, -1, 0, 0, None),
        ]

        for row, col, row_abs, col_abs, range in tests:
            exp = range
            got = xl_rowcol_to_cell(row, col, row_abs, col_abs)

            # Ignore the warnings for negative values.
            warnings.filterwarnings("ignore")

            self.assertEqual(got, exp)

    def test_xl_rowcol_to_cell_fast(self):
        """Test xl_rowcol_to_cell_fast()"""

        tests = [
            # row, col, A1 range
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

        for row, col, range in tests:
            exp = range
            got = xl_rowcol_to_cell_fast(row, col)
            self.assertEqual(got, exp)
