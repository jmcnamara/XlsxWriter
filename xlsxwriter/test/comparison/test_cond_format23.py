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
        self.set_filename("cond_format23.xlsx")

    def test_create_file(self):
        """Test the rule priorities when a range is repeated after a different range."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        format1 = workbook.add_format({"bg_color": "#FF0000", "dxf_index": 2})
        format2 = workbook.add_format({"bg_color": "#0000FF", "dxf_index": 0})
        format3 = workbook.add_format({"bg_color": "#008000", "dxf_index": 1})

        for row in range(5):
            worksheet.write(row, 0, 1)
            worksheet.write(row, 1, row)

        worksheet.conditional_format(
            "A1:A5",
            {"type": "cell", "criteria": ">", "value": 5, "format": format1},
        )

        worksheet.conditional_format(
            "A1:B5",
            {"type": "cell", "criteria": ">", "value": 0, "format": format2},
        )

        worksheet.conditional_format(
            "A1:A5",
            {"type": "cell", "criteria": ">", "value": 0, "format": format3},
        )

        workbook.close()

        self.assertExcelEqual()
