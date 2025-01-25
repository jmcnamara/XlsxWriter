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
        self.set_filename("background06.xlsx")

        self.ignore_elements = {"xl/worksheets/sheet1.xml": ["<pageSetup"]}

    def test_create_file(self):
        """Test the creation of an XlsxWriter file with a background image."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.insert_image("E9", self.image_dir + "logo.jpg")
        worksheet.set_background(self.image_dir + "logo.jpg")

        worksheet.set_header("&C&G", {"image_center": self.image_dir + "blue.jpg"})

        workbook.close()

        self.assertExcelEqual()
