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
        self.set_filename("quote_name01.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        data = [
            [1, 2, 3, 4, 5],
            [2, 4, 6, 8, 10],
            [3, 6, 9, 12, 15],
        ]

        # Test quoted/non-quoted sheet names.
        worksheet = workbook.add_worksheet("Sheet 1")
        chart = workbook.add_chart({"type": "pie"})

        worksheet.write_column("A1", data[0])
        worksheet.write_column("B1", data[1])
        worksheet.write_column("C1", data[2])

        chart.add_series({"values": ["'Sheet 1'", 0, 0, 4, 0]})
        worksheet.insert_chart("E6", chart, {"x_offset": 26, "y_offset": 17})

        sheetnames = (
            "Sheet 2",
            "Sheet!3",
            'Sheet"4',
            "Sheet#5",
            "Sheet$6",
            "Sheet%7",
            "Sheet&8",
        )

        for sheetname in sheetnames:
            worksheet = workbook.add_worksheet(sheetname)
            chart = workbook.add_chart({"type": "pie"})

            worksheet.write_column("A1", data[0])
            worksheet.write_column("B1", data[1])
            worksheet.write_column("C1", data[2])

            chart.add_series({"values": [sheetname, 0, 0, 4, 0]})
            worksheet.insert_chart("E6", chart, {"x_offset": 26, "y_offset": 17})

        workbook.close()

        self.assertExcelEqual()
