###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

from xlsxwriter.color import Color
from xlsxwriter.workbook import Workbook

from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("theme_color04.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with a theme color."""
        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        # Add theme colors to the worksheet.
        for row in range(6):
            col = 0
            color = col + 3  # Theme color index.
            shade = row + 0  # Theme shade index.
            theme_color = Color((color, shade))
            color_format = workbook.add_format({"bg_color": theme_color})

            worksheet.write(row, col, "", color_format)

        workbook.close()

        self.assertExcelEqual()
