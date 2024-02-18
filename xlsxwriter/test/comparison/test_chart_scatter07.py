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
        self.set_filename("chart_scatter07.xlsx")

        self.ignore_elements = {"xl/workbook.xml": ["<fileVersion", "<calcPr"]}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({"type": "scatter"})

        chart.axis_ids = [63597952, 63616128]
        chart.axis2_ids = [63617664, 63619456]

        data = [
            [27, 33, 44, 12, 1],
            [6, 8, 6, 4, 2],
            [20, 10, 30, 50, 40],
            [0, 27, 23, 30, 40],
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
                "categories": "=Sheet1!$C$1:$C$5",
                "values": "=Sheet1!$D$1:$D$5",
                "y2_axis": 1,
            }
        )

        worksheet.insert_chart("E9", chart)

        workbook.close()

        self.assertExcelEqual()
