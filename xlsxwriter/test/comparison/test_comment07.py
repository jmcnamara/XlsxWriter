###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

from xlsxwriter.workbook import Workbook

from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("comment07.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with comments."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.write_comment("A1", "Some text")
        worksheet.write_comment("A2", "Some text")
        worksheet.write_comment("A3", "Some text")
        worksheet.write_comment("A4", "Some text")
        worksheet.write_comment("A5", "Some text")

        worksheet.show_comments()

        worksheet.set_comments_author("John")

        workbook.close()

        self.assertExcelEqual()
