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
        self.set_filename("chart_gridlines09.xlsx")

    def test_create_file(self):
        """Test XlsxWriter gridlines."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({"type": "column"})

        chart.axis_ids = [48744320, 49566848]

        data = [
            [1, 2, 3, 4, 5],
            [2, 4, 6, 8, 10],
            [3, 6, 9, 12, 15],
        ]

        worksheet.write_column("A1", data[0])
        worksheet.write_column("B1", data[1])
        worksheet.write_column("C1", data[2])

        chart.add_series({"values": "=Sheet1!$A$1:$A$5"})
        chart.add_series({"values": "=Sheet1!$B$1:$B$5"})
        chart.add_series({"values": "=Sheet1!$C$1:$C$5"})

        chart.set_x_axis(
            {
                "major_gridlines": {
                    "visible": 1,
                    "line": {"color": "red", "width": 0.5, "dash_type": "square_dot"},
                },
                "minor_gridlines": {"visible": 1, "line": {"color": "yellow"}},
            }
        )

        chart.set_y_axis(
            {
                "major_gridlines": {
                    "visible": 1,
                    "line": {"width": 1.25, "dash_type": "dash"},
                },
                "minor_gridlines": {"visible": 1, "line": {"color": "#00B050"}},
            }
        )

        worksheet.insert_chart("E9", chart)

        workbook.close()

        self.assertExcelEqual()
