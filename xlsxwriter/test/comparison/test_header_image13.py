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
        self.set_filename("header_image13.xlsx")

        self.ignore_elements = {
            "xl/worksheets/sheet1.xml": ["<pageMargins", "<pageSetup"]
        }

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_header(
            "&L&G&C&G&R&G",
            {
                "image_left": self.image_dir + "black_72.jpg",
                "image_center": self.image_dir + "black_150.jpg",
                "image_right": self.image_dir + "black_300.jpg",
            },
        )

        workbook.close()

        self.assertExcelEqual()
