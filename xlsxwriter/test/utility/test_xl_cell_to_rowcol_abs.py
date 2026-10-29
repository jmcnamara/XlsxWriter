###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import unittest

from xlsxwriter.utility import xl_cell_to_rowcol_abs


class TestUtility(unittest.TestCase):
    """
    Test xl_cell_to_rowcol_abs() utility function.

    """

    def test_xl_cell_to_rowcol_abs(self):
        """Test xl_cell_to_rowcol_abs()"""

        self.assertEqual((0, 0, 0, 0), xl_cell_to_rowcol_abs("A1"))
        self.assertEqual((0, 1, 0, 0), xl_cell_to_rowcol_abs("B1"))
        self.assertEqual((0, 2, 0, 0), xl_cell_to_rowcol_abs("C1"))
        self.assertEqual((0, 9, 0, 0), xl_cell_to_rowcol_abs("J1"))
        self.assertEqual((1, 0, 0, 0), xl_cell_to_rowcol_abs("A2"))
        self.assertEqual((2, 0, 0, 0), xl_cell_to_rowcol_abs("A3"))
        self.assertEqual((9, 0, 0, 0), xl_cell_to_rowcol_abs("A10"))
        self.assertEqual((1, 24, 0, 0), xl_cell_to_rowcol_abs("Y2"))
        self.assertEqual((7, 25, 0, 0), xl_cell_to_rowcol_abs("Z8"))
        self.assertEqual((9, 26, 0, 0), xl_cell_to_rowcol_abs("AA10"))
        self.assertEqual((1, 254, 0, 0), xl_cell_to_rowcol_abs("IU2"))
        self.assertEqual((1, 255, 0, 0), xl_cell_to_rowcol_abs("IV2"))
        self.assertEqual((1, 256, 0, 0), xl_cell_to_rowcol_abs("IW2"))
        self.assertEqual((0, 16383, 0, 0), xl_cell_to_rowcol_abs("XFD1"))
        self.assertEqual((1048576, 16384, 0, 0), xl_cell_to_rowcol_abs("XFE1048577"))

    def test_xl_cell_to_rowcol_abs_abs(self):
        """Test xl_cell_to_rowcol_abs() with absolute references"""

        self.assertEqual((0, 0, 0, 0), xl_cell_to_rowcol_abs("A1"))
        self.assertEqual((0, 0, 1, 0), xl_cell_to_rowcol_abs("A$1"))
        self.assertEqual((0, 0, 0, 1), xl_cell_to_rowcol_abs("$A1"))
        self.assertEqual((0, 0, 1, 1), xl_cell_to_rowcol_abs("$A$1"))
