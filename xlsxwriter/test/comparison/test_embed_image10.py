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

        self.set_filename("embed_image10.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.embed_image(
            0, 0, self.image_dir + "red.png", {"url": "http://www.cpan.org/"}
        )

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_with_image_and_url_objects(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        image = Image(self.image_dir + "red.png")
        image.url = Url("http://www.cpan.org/")

        worksheet.embed_image(0, 0, image)

        workbook.close()

        self.assertExcelEqual()
