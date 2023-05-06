###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2023, John McNamara, jmcnamara@cpan.org
#

from io import BytesIO

from ...workbook import Workbook
from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("background02.xlsx")

    def test_create_file(self):
        """Test the creation of an XlsxWriter file with a background image."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_background(self.image_dir + "logo.jpg")

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_bytestream(self):
        """Test the creation of an XlsxWriter file with a background image."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        image_file = open(self.image_dir + "logo.jpg", "rb")
        image_data = BytesIO(image_file.read())
        image_file.close()

        worksheet.set_background(image_data, is_byte_stream=True)

        workbook.close()

        self.assertExcelEqual()
