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
        self.set_filename("chart_points03.xlsx")

    def test_create_file(self):
        """Test the creation of an XlsxWriter file with point formatting."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({"type": "pie"})

        workbook.custom_colors = ["FFCC0000", "FF990000"]

        data = [2, 5, 4]

        worksheet.write_column("A1", data)

        chart.add_series(
            {
                "values": "=Sheet1!$A$1:$A$3",
                "points": [
                    {"fill": {"color": "#FF0000"}},
                    {"fill": {"color": "#CC0000"}},
                    {"fill": {"color": "#990000"}},
                ],
            }
        )

        worksheet.insert_chart("E9", chart)

        workbook.close()

        self.assertExcelEqual()
