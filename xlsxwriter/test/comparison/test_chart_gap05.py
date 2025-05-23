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
        self.set_filename("chart_gap05.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({"type": "bar"})

        chart.axis_ids = [45938176, 59715584]
        chart.axis2_ids = [70848512, 54519680]

        data = [[1, 2, 3, 4, 5], [6, 8, 6, 4, 2]]

        worksheet.write_column("A1", data[0])
        worksheet.write_column("B1", data[1])

        chart.add_series({"values": "=Sheet1!$A$1:$A$5", "gap": 51, "overlap": 12})
        chart.add_series(
            {"values": "=Sheet1!$B$1:$B$5", "y2_axis": 1, "gap": 251, "overlap": -27}
        )

        chart.set_x2_axis({"label_position": "next_to"})

        worksheet.insert_chart("E9", chart)

        workbook.close()

        self.assertExcelEqual()
