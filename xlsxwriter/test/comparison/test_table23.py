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
        self.set_filename("table23.xlsx")

        self.ignore_files = [
            "xl/calcChain.xml",
            "[Content_Types].xml",
            "xl/_rels/workbook.xml.rels",
        ]

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with tables."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_column("B:F", 10.288)

        worksheet.write_string("A1", "Column1")
        worksheet.write_string("F1", "Total")
        worksheet.write_string("B1", "Column'")
        worksheet.write_string("C1", "Column#")
        worksheet.write_string("D1", "Column[")
        worksheet.write_string("E1", "Column]")

        worksheet.add_table(
            "B3:F9",
            {
                "total_row": True,
                "columns": [
                    {"header": "Column1", "total_string": "Total"},
                    {"header": "Column'", "total_function": "sum"},
                    {"header": "Column#", "total_function": "sum"},
                    {"header": "Column[", "total_function": "sum"},
                    {"header": "Column]", "total_function": "sum"},
                ],
            },
        )

        workbook.close()

        self.assertExcelEqual()
