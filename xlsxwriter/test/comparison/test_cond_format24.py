###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

from xlsxwriter.workbook import Workbook

from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("cond_format24.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with conditionalFormatting."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        format1 = workbook.add_format({"bg_color": "red"})
        format2 = workbook.add_format({"bg_color": "green"})

        worksheet.write_column("A1", [1, 1, 1, 1, 1])
        worksheet.write_column("B1", [1, 2, 3, 4, 5])

        worksheet.conditional_format(
            "A1:A5",
            {"type": "cell", "criteria": ">=", "value": 0, "format": format1},
        )

        # The 5_boxes icon set is only available in Excel 2010 and later.
        worksheet.conditional_format(
            "B1:B5", {"type": "icon_set", "icon_style": "5_boxes"}
        )

        worksheet.conditional_format(
            "A1:A5",
            {"type": "cell", "criteria": ">", "value": 5, "format": format2},
        )

        workbook.close()

        self.assertExcelEqual()
