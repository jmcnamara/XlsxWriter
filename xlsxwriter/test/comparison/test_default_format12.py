###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

from xlsxwriter.workbook import Workbook

from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("default_format12.xlsx")

    def test_create_file(self):
        """Test the creation of a file with user defined default format"""

        workbook = Workbook(
            self.got_filename,
            {
                "default_format_properties": {"font_name": "Arial", "font_size": 14},
                "default_row_height": 24,
                "default_column_width": 96,
            },
        )

        worksheet = workbook.add_worksheet()

        worksheet.insert_image("E9", self.image_dir + "red.png", {"x_offset": 32})

        # Set user column width and row height to test positioning calculation.
        # The column width is the default in this font.
        worksheet.set_row_pixels(8, 32)

        # Set column to text column width less than 1 character.
        worksheet.set_column_pixels(6, 6, 10)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_with_character_units(self):
        """Test the creation of a file with user defined default format"""

        # Same as
        workbook = Workbook(
            self.got_filename,
            {
                "default_format_properties": {"font_name": "Arial", "font_size": 14},
                "default_row_height": 24,
                "default_column_width": 96,
            },
        )

        worksheet = workbook.add_worksheet()

        worksheet.insert_image("E9", self.image_dir + "red.png", {"x_offset": 32})

        # Set user column width and row height to test positioning calculation.
        # The column width is the default in this font.
        worksheet.set_row(8, 24.0)

        # Set column to text column width less than 1 character.
        worksheet.set_column(6, 6, 0.56)

        workbook.close()

        self.assertExcelEqual()
