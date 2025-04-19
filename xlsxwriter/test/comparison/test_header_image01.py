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
        self.set_filename("header_image01.xlsx")

        self.ignore_elements = {
            "xl/worksheets/sheet1.xml": ["<pageMargins", "<pageSetup"]
        }

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_header("&L&G", {"image_left": self.image_dir + "red.jpg"})

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_in_memory(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename, {"in_memory": True})

        worksheet = workbook.add_worksheet()

        worksheet.set_header("&L&G", {"image_left": self.image_dir + "red.jpg"})

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_from_bytesio(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        with open(self.image_dir + "red.jpg", "rb") as image_file:
            image_data = BytesIO(image_file.read())

        worksheet.set_header(
            "&L&G", {"image_left": "red.jpg", "image_data_left": image_data}
        )

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_from_bytesio_in_memory(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename, {"in_memory": True})

        worksheet = workbook.add_worksheet()

        with open(self.image_dir + "red.jpg", "rb") as image_file:
            image_data = BytesIO(image_file.read())

        worksheet.set_header(
            "&L&G", {"image_left": "red.jpg", "image_data_left": image_data}
        )

        workbook.close()

        self.assertExcelEqual()
