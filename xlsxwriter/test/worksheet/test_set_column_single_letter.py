###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import BytesIO

from xlsxwriter.workbook import Workbook


class TestSetColumnSingleLetter(unittest.TestCase):
    """Single-column A1 letters must work as A:A, not ValueError unpack."""

    def test_set_column_single_letter(self):
        workbook = Workbook(BytesIO(), {"in_memory": True})
        worksheet = workbook.add_worksheet()
        self.assertEqual(worksheet.set_column("A", 10), 0)
        self.assertEqual(worksheet.set_column("AA", 12), 0)
        # Same pixel width as set_column("A:A", 10) / ("AA:AA", 12).
        self.assertEqual(worksheet.col_info[0].width, 75)
        self.assertEqual(worksheet.col_info[26].width, 89)
        workbook.close()

    def test_set_column_a1_range_still_ok(self):
        workbook = Workbook(BytesIO(), {"in_memory": True})
        worksheet = workbook.add_worksheet()
        self.assertEqual(worksheet.set_column("B:D", 5), 0)
        workbook.close()
