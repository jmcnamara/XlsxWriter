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
        self.set_filename("chart_errorbars10.xlsx")

        # Test for issue #115. We don't add plus_data and minus_data, like in
        # test_chart_errorbars10.py, as would be done from user API. Instead
        # we ignore the point data in the comparison test.
        self.ignore_elements = {
            "xl/charts/chart1.xml": ["<c:ptCount", "<c:pt", "<c:v", "</c:pt>"]
        }

    def test_create_file(self):
        """Test the creation of an XlsxWriter file with error bars."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({"type": "line"})

        chart.axis_ids = [69198976, 69200896]

        data = [
            [1, 2, 3, 4, 5],
            [2, 4, 6, 8, 10],
            [3, 6, 9, 12, 15],
        ]

        worksheet.write_column("A1", data[0])
        worksheet.write_column("B1", data[1])
        worksheet.write_column("C1", data[2])

        chart.add_series(
            {
                "categories": "=Sheet1!$A$1:$A$5",
                "values": "=Sheet1!$B$1:$B$5",
                "y_error_bars": {
                    "type": "custom",
                    "plus_values": "=Sheet1!$A$1",
                    "minus_values": "=Sheet1!$B$1:$B$3",
                },
            }
        )

        chart.add_series(
            {
                "categories": "=Sheet1!$A$1:$A$5",
                "values": "=Sheet1!$C$1:$C$5",
            }
        )

        worksheet.insert_chart("E9", chart)

        workbook.close()

        self.assertExcelEqual()
