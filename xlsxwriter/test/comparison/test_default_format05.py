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
        self.set_filename("default_format05.xlsx")

    def test_create_file(self):
        """Test the creation of a file with user defined default format"""

        workbook = Workbook(
            self.got_filename,
            {
                "default_format_properties": {
                    "font_name": "MS Gothic",
                    "font_size": 11,
                },
                "default_row_height": 18,
                "default_column_width": 72,
            },
        )

        worksheet = workbook.add_worksheet()

        worksheet.insert_image("E9", self.image_dir + "red.png")

        workbook.close()

        self.assertExcelEqual()
