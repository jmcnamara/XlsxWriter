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
        self.set_filename("chart_chartarea03.xlsx")

    def test_create_file(self):
        """Test XlsxWriter chartarea properties."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({"type": "scatter"})

        chart.axis_ids = [46210048, 46208512]

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
            }
        )

        chart.add_series(
            {
                "categories": "=Sheet1!$A$1:$A$5",
                "values": "=Sheet1!$C$1:$C$5",
            }
        )

        chart.set_chartarea(
            {"border": {"dash_type": "round_dot"}, "fill": {"color": "#9999FF"}}
        )

        chart.set_plotarea(
            {"border": {"dash_type": "square_dot"}, "fill": {"color": "#FFC000"}}
        )

        worksheet.insert_chart("E9", chart)

        workbook.close()

        self.assertExcelEqual()
