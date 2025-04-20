###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

from xlsxwriter.color import Color
from xlsxwriter.workbook import Workbook

from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("chart_theme05.xlsx")

    def test_create_file(self):
        """Test the creation of an XlsxWriter file with chart formatting."""
        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        chart = workbook.add_chart({"type": "line", "subtype": "stacked"})
        chart.axis_ids = [68411392, 68414848]

        # Add some test data for the chart(s).
        for row_num in range(8):
            for col_num in range(6):
                worksheet.write_number(row_num, col_num, 1)

        chart.add_series(
            {
                "values": ["Sheet1", 0, 0, 7, 0],
                "line": {"color": Color((3, 0))},
            }
        )
        chart.add_series(
            {
                "values": ["Sheet1", 0, 1, 7, 1],
                "line": {"color": Color((3, 1))},
            }
        )
        chart.add_series(
            {
                "values": ["Sheet1", 0, 2, 7, 2],
                "line": {"color": Color((3, 2))},
            }
        )
        chart.add_series(
            {
                "values": ["Sheet1", 0, 3, 7, 3],
                "line": {"color": Color((3, 3))},
            }
        )
        chart.add_series(
            {
                "values": ["Sheet1", 0, 4, 7, 4],
                "line": {"color": Color((3, 4))},
            }
        )
        chart.add_series(
            {
                "values": ["Sheet1", 0, 5, 7, 5],
                "line": {"color": Color((3, 5))},
            }
        )

        worksheet.insert_chart(8, 7, chart)

        workbook.close()

        self.assertExcelEqual()
