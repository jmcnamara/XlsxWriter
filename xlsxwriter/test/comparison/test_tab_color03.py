###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

from xlsxwriter.color import Color
from xlsxwriter.workbook import Workbook

from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("tab_color03.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with a tab color."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.vertical_dpi = 200

        worksheet.fit_to_pages(1, 1)
        worksheet.outline_settings(symbols_below=False)

        worksheet.write("A1", "Foo")
        worksheet.set_tab_color(Color("red"))

        workbook.close()

        self.assertExcelEqual()
