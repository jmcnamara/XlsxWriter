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
        self.set_filename("cond_format12.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with conditional formatting."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        format1 = workbook.add_format(
            {"bg_color": "#FFFF00", "fg_color": "#FF0000", "pattern": 12}
        )

        worksheet.write("A1", "Hello", format1)

        worksheet.write("B3", 10)
        worksheet.write("B4", 20)
        worksheet.write("B5", 30)
        worksheet.write("B6", 40)

        worksheet.conditional_format(
            "B3:B6",
            {
                "type": "cell",
                "format": format1,
                "criteria": "greater than",
                "value": 20,
            },
        )

        workbook.close()

        self.assertExcelEqual()
