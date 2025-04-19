###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

from io import BytesIO

from xlsxwriter.workbook import Workbook

from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("header_image04.xlsx")

        self.ignore_elements = {
            "xl/worksheets/sheet1.xml": ["<pageMargins", "<pageSetup"]
        }

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_footer(
            "&L&G&C&G&R&G",
            {
                "image_left": self.image_dir + "red.jpg",
                "image_center": self.image_dir + "blue.jpg",
                "image_right": self.image_dir + "yellow.jpg",
            },
        )

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_with_picture(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_footer(
            "&L&[Picture]&C&G&R&[Picture]",
            {
                "image_left": self.image_dir + "red.jpg",
                "image_center": self.image_dir + "blue.jpg",
                "image_right": self.image_dir + "yellow.jpg",
            },
        )

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_from_bytesio(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        with open(self.image_dir + "red.jpg", "rb") as image_file_left:
            image_data_left = BytesIO(image_file_left.read())

        with open(self.image_dir + "blue.jpg", "rb") as image_file_center:
            image_data_center = BytesIO(image_file_center.read())

        with open(self.image_dir + "yellow.jpg", "rb") as image_file_right:
            image_data_right = BytesIO(image_file_right.read())

        worksheet.set_footer(
            "&L&G&C&G&R&G",
            {
                "image_left": "red.jpg",
                "image_center": "blue.jpg",
                "image_right": "yellow.jpg",
                "image_data_left": image_data_left,
                "image_data_center": image_data_center,
                "image_data_right": image_data_right,
            },
        )

        workbook.close()

        self.assertExcelEqual()
