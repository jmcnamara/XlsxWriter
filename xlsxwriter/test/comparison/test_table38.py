###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

from ...workbook import Workbook
from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):

        self.set_filename("table38.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with tables."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.write(1, 0, 1)
        worksheet.write(2, 0, 2)
        worksheet.write(3, 0, 3)
        worksheet.write(4, 0, 4)
        worksheet.write(5, 0, 5)

        worksheet.write(1, 1, 10)
        worksheet.write(2, 1, 15)
        worksheet.write(3, 1, 20)
        worksheet.write(4, 1, 10)
        worksheet.write(5, 1, 15)

        worksheet.set_column("A:B", 10.288)

        worksheet.add_table("A1:B6", {"description": "Alt text", "title": "Alt title"})

        workbook.close()

        self.assertExcelEqual()
