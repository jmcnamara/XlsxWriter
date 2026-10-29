###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import unittest
import warnings

from xlsxwriter.utility import xl_rowcol_to_cell, xl_rowcol_to_cell_fast


class TestUtility(unittest.TestCase):
    """
    Test xl_rowcol_to_cell() utility function.

    """

    def test_xl_rowcol_to_cell(self):
        """Test xl_rowcol_to_cell()"""

        self.assertEqual("A1", xl_rowcol_to_cell(0, 0))
        self.assertEqual("B1", xl_rowcol_to_cell(0, 1))
        self.assertEqual("C1", xl_rowcol_to_cell(0, 2))
        self.assertEqual("J1", xl_rowcol_to_cell(0, 9))
        self.assertEqual("A2", xl_rowcol_to_cell(1, 0))
        self.assertEqual("A3", xl_rowcol_to_cell(2, 0))
        self.assertEqual("A10", xl_rowcol_to_cell(9, 0))
        self.assertEqual("Y2", xl_rowcol_to_cell(1, 24))
        self.assertEqual("Z8", xl_rowcol_to_cell(7, 25))
        self.assertEqual("AA10", xl_rowcol_to_cell(9, 26))
        self.assertEqual("IU2", xl_rowcol_to_cell(1, 254))
        self.assertEqual("IV2", xl_rowcol_to_cell(1, 255))
        self.assertEqual("IW2", xl_rowcol_to_cell(1, 256))
        self.assertEqual("XFD1", xl_rowcol_to_cell(0, 16383))
        self.assertEqual("XFE1048577", xl_rowcol_to_cell(1048576, 16384))

        # Ignore the warnings for negative values.
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            self.assertEqual("", xl_rowcol_to_cell(0, -1))
            self.assertEqual("", xl_rowcol_to_cell(-1, 0))

    def test_xl_rowcol_to_cell_abs(self):
        """Test xl_rowcol_to_cell() with absolute references"""

        self.assertEqual("A$1", xl_rowcol_to_cell(0, 0, True, False))
        self.assertEqual("$A1", xl_rowcol_to_cell(0, 0, False, True))
        self.assertEqual("$A$1", xl_rowcol_to_cell(0, 0, True, True))

        # Ignore the warnings for negative values.
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            self.assertEqual("", xl_rowcol_to_cell(-1, 0, 0, 0))
            self.assertEqual("", xl_rowcol_to_cell(0, -1, 0, 0))

    def test_xl_rowcol_to_cell_fast(self):
        """Test xl_rowcol_to_cell_fast()"""

        self.assertEqual("A1", xl_rowcol_to_cell_fast(0, 0))
        self.assertEqual("B1", xl_rowcol_to_cell_fast(0, 1))
        self.assertEqual("C1", xl_rowcol_to_cell_fast(0, 2))
        self.assertEqual("J1", xl_rowcol_to_cell_fast(0, 9))
        self.assertEqual("A2", xl_rowcol_to_cell_fast(1, 0))
        self.assertEqual("A3", xl_rowcol_to_cell_fast(2, 0))
        self.assertEqual("A10", xl_rowcol_to_cell_fast(9, 0))
        self.assertEqual("Y2", xl_rowcol_to_cell_fast(1, 24))
        self.assertEqual("Z8", xl_rowcol_to_cell_fast(7, 25))
        self.assertEqual("AA10", xl_rowcol_to_cell_fast(9, 26))
        self.assertEqual("IU2", xl_rowcol_to_cell_fast(1, 254))
        self.assertEqual("IV2", xl_rowcol_to_cell_fast(1, 255))
        self.assertEqual("IW2", xl_rowcol_to_cell_fast(1, 256))
        self.assertEqual("XFD1", xl_rowcol_to_cell_fast(0, 16383))
        self.assertEqual("XFE1048577", xl_rowcol_to_cell_fast(1048576, 16384))
