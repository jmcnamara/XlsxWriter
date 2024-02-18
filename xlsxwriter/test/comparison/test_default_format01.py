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
        self.set_filename("default_format01.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(
            self.got_filename, {"default_format_properties": {"font_size": 10}}
        )

        worksheet = workbook.add_worksheet()

        worksheet.set_default_row(12.75)

        # For testing.
        worksheet.original_row_height = 12.75

        workbook.close()

        self.assertExcelEqual()
