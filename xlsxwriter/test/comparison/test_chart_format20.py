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
        self.set_filename("chart_format20.xlsx")

    def test_create_file(self):
        """Test the creation of an XlsxWriter file with chart formatting."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart1 = workbook.add_chart({"type": "line"})
        chart2 = workbook.add_chart({"type": "line"})

        chart1.axis_ids = [80553856, 80555392]
        chart2.axis_ids = [84583936, 84585856]

        trend = {"type": "linear", "line": {"color": "red", "dash_type": "dash"}}

        data = [
            [1, 2, 3, 4, 5],
            [2, 4, 6, 8, 10],
            [3, 6, 9, 12, 15],
        ]

        worksheet.write_column("A1", data[0])
        worksheet.write_column("B1", data[1])
        worksheet.write_column("C1", data[2])

        chart1.add_series(
            {
                "values": "=Sheet1!$B$1:$B$5",
                "trendline": trend,
            }
        )

        chart1.add_series(
            {
                "values": "=Sheet1!$C$1:$C$5",
            }
        )

        chart2.add_series(
            {
                "values": "=Sheet1!$B$1:$B$5",
                "trendline": trend,
            }
        )

        chart2.add_series(
            {
                "values": "=Sheet1!$C$1:$C$5",
            }
        )

        worksheet.insert_chart("E9", chart1)
        worksheet.insert_chart("E25", chart2)

        workbook.close()

        self.assertExcelEqual()
