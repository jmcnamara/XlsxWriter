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
        self.set_filename("chart_display_units12.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({"type": "scatter"})

        chart.axis_ids = [93550464, 93548544]

        data = [
            [10000000, 20000000, 30000000, 20000000, 10000000],
        ]

        worksheet.write_column(0, 0, data[0])
        worksheet.write_column(0, 1, data[0])

        chart.add_series(
            {"categories": "=Sheet1!$A$1:$A$5", "values": "=Sheet1!$B$1:$B$5"}
        )

        chart.set_y_axis({"display_units": "hundreds", "display_units_visible": False})
        chart.set_x_axis({"display_units": "thousands", "display_units_visible": False})

        worksheet.insert_chart("E9", chart)

        workbook.close()

        self.assertExcelEqual()
