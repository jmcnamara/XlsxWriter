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
        self.set_filename("chart_errorbars07.xlsx")

        self.ignore_elements = {"xl/charts/chart1.xml": ["<c:formatCode"]}

    def test_create_file(self):
        """Test the creation of an XlsxWriter file with error bars."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({"type": "stock"})
        date_format = workbook.add_format({"num_format": 14})

        chart.axis_ids = [45470848, 45472768]

        data = [
            [39083, 39084, 39085, 39086, 39087],
            [27.2, 25.03, 19.05, 20.34, 18.5],
            [23.49, 19.55, 15.12, 17.84, 16.34],
            [25.45, 23.05, 17.32, 20.45, 17.34],
        ]

        for row in range(5):
            worksheet.write(row, 0, data[0][row], date_format)
            worksheet.write(row, 1, data[1][row])
            worksheet.write(row, 2, data[2][row])
            worksheet.write(row, 3, data[3][row])

        worksheet.set_column("A:D", 11)

        chart.add_series(
            {
                "categories": "=Sheet1!$A$1:$A$5",
                "values": "=Sheet1!$B$1:$B$5",
                "y_error_bars": {"type": "standard_error"},
            }
        )

        chart.add_series(
            {
                "categories": "=Sheet1!$A$1:$A$5",
                "values": "=Sheet1!$C$1:$C$5",
                "y_error_bars": {"type": "standard_error"},
            }
        )

        chart.add_series(
            {
                "categories": "=Sheet1!$A$1:$A$5",
                "values": "=Sheet1!$D$1:$D$5",
                "y_error_bars": {"type": "standard_error"},
            }
        )

        worksheet.insert_chart("E9", chart)

        workbook.close()

        self.assertExcelEqual()
