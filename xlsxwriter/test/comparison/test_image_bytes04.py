###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

from io import BytesIO

from ...workbook import Workbook
from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("image48.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename)

        worksheet1 = workbook.add_worksheet()
        worksheet2 = workbook.add_worksheet()

        with open(self.image_dir + "red.png", "rb") as image_file:
            image_data = BytesIO(image_file.read())

        worksheet1.insert_image("E9", "red.png", {"image_data": image_data})
        worksheet2.insert_image("E9", "red.png", {"image_data": image_data})

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_in_memory(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename, {"in_memory": True})

        worksheet1 = workbook.add_worksheet()
        worksheet2 = workbook.add_worksheet()

        with open(self.image_dir + "red.png", "rb") as image_file:
            image_data = BytesIO(image_file.read())

        worksheet1.insert_image("E9", "red.png", {"image_data": image_data})
        worksheet2.insert_image("E9", "red.png", {"image_data": image_data})

        workbook.close()

        self.assertExcelEqual()
