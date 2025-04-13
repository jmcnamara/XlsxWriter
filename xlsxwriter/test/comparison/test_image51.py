###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

from ...image import Image
from ...url import Url
from ...workbook import Workbook
from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("image51.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.insert_image(
            "E9", self.image_dir + "red.png", {"url": "https://duckduckgo.com/?q=1"}
        )
        worksheet.insert_image(
            "E13", self.image_dir + "red2.png", {"url": "https://duckduckgo.com/?q=2"}
        )

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_with_image_and_url_objects(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        url1 = Url("https://duckduckgo.com/?q=1")
        url2 = Url("https://duckduckgo.com/?q=2")

        image1 = Image(self.image_dir + "red.png")
        image2 = Image(self.image_dir + "red2.png")

        image1.url = url1
        image2.url = url2

        worksheet.insert_image("E9", image1)
        worksheet.insert_image("E13", image2)

        workbook.close()

        self.assertExcelEqual()
