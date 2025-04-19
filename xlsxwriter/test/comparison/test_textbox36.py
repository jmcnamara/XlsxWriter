###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

from xlsxwriter.url import Url
from xlsxwriter.workbook import Workbook

from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("textbox36.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with textbox(s)."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        worksheet.insert_textbox(
            "E9", "This is some text", {"url": "https://github.com/jmcnamara"}
        )

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_with_url_object(self):
        """Test the creation of a simple XlsxWriter file with textbox(s)."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        url = Url("https://github.com/jmcnamara")

        worksheet.insert_textbox("E9", "This is some text", {"url": url})

        workbook.close()

        self.assertExcelEqual()
