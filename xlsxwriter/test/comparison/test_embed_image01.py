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

        self.set_filename("embed_image01.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.embed_image(0, 0, self.image_dir + "red.png")

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_from_memory(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        with open(self.image_dir + "red.png", "rb") as image_file:
            image_data = BytesIO(image_file.read())

        worksheet.embed_image(0, 0, "", {"image_data": image_data})

        workbook.close()

        self.assertExcelEqual()
