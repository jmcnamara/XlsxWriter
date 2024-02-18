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
        self.set_filename("image49.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename)

        worksheet1 = workbook.add_worksheet()
        worksheet2 = workbook.add_worksheet()
        worksheet3 = workbook.add_worksheet()

        worksheet1.insert_image("A1", self.image_dir + "blue.png")
        worksheet1.insert_image("B3", self.image_dir + "red.jpg")
        worksheet1.insert_image("D5", self.image_dir + "yellow.jpg")
        worksheet1.insert_image("F9", self.image_dir + "grey.png")

        worksheet2.insert_image("A1", self.image_dir + "blue.png")
        worksheet2.insert_image("B3", self.image_dir + "red.jpg")
        worksheet2.insert_image("D5", self.image_dir + "yellow.jpg")
        worksheet2.insert_image("F9", self.image_dir + "grey.png")

        worksheet3.insert_image("A1", self.image_dir + "blue.png")
        worksheet3.insert_image("B3", self.image_dir + "red.jpg")
        worksheet3.insert_image("D5", self.image_dir + "yellow.jpg")
        worksheet3.insert_image("F9", self.image_dir + "grey.png")

        workbook.close()

        self.assertExcelEqual()
