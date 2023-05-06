###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2023, John McNamara, jmcnamara@cpan.org
#

from ...workbook import Workbook
from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("cond_format08.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with conditional formatting."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        format = workbook.add_format(
            {
                "color": "#9C6500",
                "bg_color": "#FFEB9C",
                "font_condense": 1,
                "font_extend": 1,
            }
        )

        worksheet.write("A1", 10)
        worksheet.write("A2", 20)
        worksheet.write("A3", 30)
        worksheet.write("A4", 40)

        worksheet.conditional_format(
            "A1",
            {"type": "cell", "format": format, "criteria": "greater than", "value": 5},
        )

        workbook.close()

        self.assertExcelEqual()
