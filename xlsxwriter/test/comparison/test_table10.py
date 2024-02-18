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
        self.set_filename("table10.xlsx")

        self.ignore_files = [
            "xl/calcChain.xml",
            "[Content_Types].xml",
            "xl/_rels/workbook.xml.rels",
        ]

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with tables."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        xformat = workbook.add_format({"num_format": 2})

        worksheet.set_column("B:K", 10.288)

        worksheet.write_string("A1", "Column1")
        worksheet.write_string("B1", "Column2")
        worksheet.write_string("C1", "Column3")
        worksheet.write_string("D1", "Column4")
        worksheet.write_string("E1", "Column5")
        worksheet.write_string("F1", "Column6")
        worksheet.write_string("G1", "Column7")
        worksheet.write_string("H1", "Column8")
        worksheet.write_string("I1", "Column9")
        worksheet.write_string("J1", "Column10")
        worksheet.write_string("K1", "Total")

        data = [0, 0, 0, None, None, 0, 0, 0, 0, 0]
        worksheet.write_row("B4", data)
        worksheet.write_row("B5", data)

        worksheet.add_table(
            "B3:K6",
            {
                "total_row": 1,
                "columns": [
                    {"total_string": "Total"},
                    {},
                    {"total_function": "average"},
                    {"total_function": "count"},
                    {"total_function": "count_nums"},
                    {"total_function": "max"},
                    {"total_function": "min"},
                    {"total_function": "sum"},
                    {"total_function": "stdDev"},
                    {
                        "total_function": "var",
                        "formula": "SUM(Table1[[#This Row],[Column1]:[Column3]])",
                        "format": xformat,
                    },
                ],
            },
        )

        workbook.close()

        self.assertExcelEqual()
