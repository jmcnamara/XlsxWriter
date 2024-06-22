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

        self.set_filename("chart_line07.xlsx")
        self.ignore_elements = {
            "xl/charts/chart1.xml": [
                "<c:crosses",
            ]
        }

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({"type": "line"})

        chart.axis_ids = [77034624, 77036544]
        chart.axis2_ids = [95388032, 103040896]

        data = [
            [1, 2, 3, 4, 5],
            [10, 40, 50, 20, 10],
            [1, 2, 3, 4, 5, 6, 7],
            [30, 10, 20, 40, 30, 10, 20],
        ]

        worksheet.write_column("A1", data[0])
        worksheet.write_column("B1", data[1])
        worksheet.write_column("C1", data[2])
        worksheet.write_column("D1", data[3])

        chart.add_series(
            {
                "categories": "=Sheet1!$A$1:$A$5",
                "values": "=Sheet1!$B$1:$B$5",
            }
        )

        chart.add_series(
            {
                "categories": "=Sheet1!$C$1:$C$7",
                "values": "=Sheet1!$D$1:$D$7",
                "y2_axis": 1,
            }
        )

        chart.set_x2_axis({"label_position": "next_to"})

        worksheet.insert_chart("E9", chart)

        workbook.close()

        self.assertExcelEqual()
