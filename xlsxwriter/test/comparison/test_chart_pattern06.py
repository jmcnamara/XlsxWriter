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
        self.set_filename("chart_pattern06.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({"type": "column"})

        chart.axis_ids = [110902272, 110756608]

        data = [
            [2, 2, 2],
            [2, 2, 2],
            [2, 2, 2],
            [2, 2, 2],
            [2, 2, 2],
            [2, 2, 2],
            [2, 2, 2],
            [2, 2, 2],
        ]

        worksheet.write_column("A1", data[0])
        worksheet.write_column("B1", data[1])
        worksheet.write_column("C1", data[2])
        worksheet.write_column("D1", data[3])
        worksheet.write_column("E1", data[4])
        worksheet.write_column("F1", data[5])
        worksheet.write_column("G1", data[6])
        worksheet.write_column("H1", data[7])

        chart.add_series(
            {
                "values": "=Sheet1!$A$1:$A$3",
                "pattern": {
                    "pattern": "percent_30",
                    "fg_color": "#C00000",
                    "bg_color": "#FFFFFF",
                },
            }
        )

        chart.add_series(
            {
                "values": "=Sheet1!$B$1:$B$3",
                "pattern": {
                    "pattern": "percent_80",
                    "fg_color": "#FF0000",
                },
            }
        )

        chart.add_series(
            {
                "values": "=Sheet1!$C$1:$C$3",
                "pattern": {
                    "pattern": "wide_downward_diagonal",
                    "fg_color": "#FFC000",
                },
            }
        )

        chart.add_series(
            {
                "values": "=Sheet1!$D$1:$D$3",
                "pattern": {
                    "pattern": "dark_vertical",
                    "fg_color": "#FFFF00",
                },
            }
        )

        chart.add_series(
            {
                "values": "=Sheet1!$E$1:$E$3",
                "pattern": {
                    "pattern": "small_confetti",
                    "fg_color": "#92D050",
                },
            }
        )

        chart.add_series(
            {
                "values": "=Sheet1!$F$1:$F$3",
                "pattern": {
                    "pattern": "weave",
                    "fg_color": "#00B050",
                },
            }
        )

        chart.add_series(
            {
                "values": "=Sheet1!$G$1:$G$3",
                "pattern": {
                    "pattern": "trellis",
                    "fg_color": "#00B0F0",
                },
            }
        )

        chart.add_series(
            {
                "values": "=Sheet1!$H$1:$H$3",
                "pattern": {
                    "pattern": "outlined_diamond",
                    "fg_color": "#0070C0",
                },
            }
        )

        worksheet.insert_chart("E9", chart)

        workbook.close()

        self.assertExcelEqual()
