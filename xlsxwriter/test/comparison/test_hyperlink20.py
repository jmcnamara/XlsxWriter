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
        self.set_filename("hyperlink20.xlsx")

    def test_hyperlink_formatting_explicit(self):
        """Test the creation of a simple XlsxWriter file with hyperlinks. This example has link formatting."""

        workbook = Workbook(self.got_filename)

        # Simulate custom color for testing.
        workbook.custom_colors = ["FF0000FF"]

        worksheet = workbook.add_worksheet()
        format1 = workbook.add_format({"color": "blue", "underline": 1})
        format2 = workbook.add_format({"color": "red", "underline": 1})

        worksheet.write_url("A1", "http://www.python.org/1", format1)
        worksheet.write_url("A2", "http://www.python.org/2", format2)

        workbook.close()

        self.assertExcelEqual()
