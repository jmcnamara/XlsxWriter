###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparison_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("set_row03.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_row_pixels(0, 1)
        worksheet.set_row_pixels(1, 2)
        worksheet.set_row_pixels(2, 3)
        worksheet.set_row_pixels(3, 4)

        worksheet.set_row_pixels(11, 12)
        worksheet.set_row_pixels(12, 13)
        worksheet.set_row_pixels(13, 14)
        worksheet.set_row_pixels(14, 15)

        worksheet.set_row_pixels(18, 19)
        worksheet.set_row_pixels(20, 21, None, {"hidden": True})
        worksheet.set_row_pixels(21, 22)

        workbook.close()

        self.assertExcelEqual()
