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
        self.set_filename("header_image22.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename)

        worksheet1 = workbook.add_worksheet()
        worksheet1.set_paper(9)
        worksheet1.vertical_dpi = 200

        worksheet1.set_header("&L&G", {"image_left": self.image_dir + "blue.png"})

        worksheet2 = workbook.add_worksheet()
        worksheet2.insert_image(0, 0, self.image_dir + "red.png")

        workbook.close()

        self.assertExcelEqual()
