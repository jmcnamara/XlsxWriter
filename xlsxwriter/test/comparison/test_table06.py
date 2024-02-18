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
        self.set_filename("table06.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with tables."""

        workbook = Workbook(self.got_filename)

        # Turn off default URL format for testing.
        workbook.default_url_format = None

        worksheet = workbook.add_worksheet()

        worksheet.set_column("C:H", 10.288)

        worksheet.add_table("C3:F13")
        worksheet.add_table("F15:H20")
        worksheet.add_table("C23:D30")

        worksheet.write("A1", "http://perl.com/")
        worksheet.write("C1", "http://perl.com/")

        worksheet.set_comments_author("John")
        worksheet.write_comment("H1", "Test1")
        worksheet.write_comment("J1", "Test2")

        worksheet.insert_image("A4", self.image_dir + "blue.png")

        workbook.close()

        self.assertExcelEqual()
