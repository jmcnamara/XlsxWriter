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
        self.set_filename("image26.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.insert_image("B2", self.image_dir + "black_72.png")
        worksheet.insert_image("B8", self.image_dir + "black_96.png")
        worksheet.insert_image("B13", self.image_dir + "black_150.png")
        worksheet.insert_image("B17", self.image_dir + "black_300.png")

        workbook.close()

        self.assertExcelEqual()
