###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import unittest

from xlsxwriter.utility import xl_cell_to_rowcol


class TestUtility(unittest.TestCase):
    """
    Test xl_cell_to_rowcol() utility function.

    """

    def test_xl_cell_to_rowcol(self):
        """Test xl_cell_to_rowcol()"""

        self.assertEqual((0, 0), xl_cell_to_rowcol("A1"))
        self.assertEqual((0, 1), xl_cell_to_rowcol("B1"))
        self.assertEqual((0, 2), xl_cell_to_rowcol("C1"))
        self.assertEqual((0, 9), xl_cell_to_rowcol("J1"))
        self.assertEqual((1, 0), xl_cell_to_rowcol("A2"))
        self.assertEqual((2, 0), xl_cell_to_rowcol("A3"))
        self.assertEqual((9, 0), xl_cell_to_rowcol("A10"))
        self.assertEqual((1, 24), xl_cell_to_rowcol("Y2"))
        self.assertEqual((7, 25), xl_cell_to_rowcol("Z8"))
        self.assertEqual((9, 26), xl_cell_to_rowcol("AA10"))
        self.assertEqual((1, 254), xl_cell_to_rowcol("IU2"))
        self.assertEqual((1, 255), xl_cell_to_rowcol("IV2"))
        self.assertEqual((1, 256), xl_cell_to_rowcol("IW2"))
        self.assertEqual((0, 16383), xl_cell_to_rowcol("XFD1"))
        self.assertEqual((1048576, 16384), xl_cell_to_rowcol("XFE1048577"))

    def test_xl_cell_to_rowcol_abs(self):
        """Test xl_cell_to_rowcol() with absolute references"""

        self.assertEqual((0, 0), xl_cell_to_rowcol("A1"))
        self.assertEqual((0, 0), xl_cell_to_rowcol("A$1"))
        self.assertEqual((0, 0), xl_cell_to_rowcol("$A1"))
        self.assertEqual((0, 0), xl_cell_to_rowcol("$A$1"))
