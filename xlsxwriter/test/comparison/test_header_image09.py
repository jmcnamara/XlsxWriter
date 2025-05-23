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
        self.set_filename("header_image09.xlsx")

        self.ignore_elements = {
            "xl/worksheets/sheet1.xml": ["<pageMargins", "<pageSetup"],
            "xl/worksheets/sheet2.xml": ["<pageMargins", "<pageSetup"],
        }

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename)

        worksheet1 = workbook.add_worksheet()
        worksheet2 = workbook.add_worksheet()

        worksheet1.write("A1", "Foo")
        worksheet1.write_comment("B2", "Some text")

        worksheet1.set_comments_author("John")

        worksheet2.set_header("&L&G", {"image_left": self.image_dir + "red.jpg"})

        workbook.close()

        self.assertExcelEqual()
