###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

from ...workbook import Workbook
from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("chart_data_labels33.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({"type": "column"})

        chart.axis_ids = [65546112, 70217728]

        data = [
            [1, 2, 3, 4, 5],
            [2, 4, 6, 8, 10],
            [3, 6, 9, 12, 15],
            [10, 20, 30, 40, 50],
        ]

        worksheet.write_column("A1", data[0])
        worksheet.write_column("B1", data[1])
        worksheet.write_column("C1", data[2])
        worksheet.write_column("D1", data[3])

        chart.add_series(
            {
                "values": "=Sheet1!$A$1:$A$5",
                "data_labels": {
                    "value": 1,
                    "custom": [{"font": {"bold": 1, "italic": 1, "baseline": -1}}],
                },
            }
        )

        chart.add_series({"values": "=Sheet1!$B$1:$B$5"})
        chart.add_series({"values": "=Sheet1!$C$1:$C$5"})

        worksheet.insert_chart("E9", chart)

        workbook.close()

        self.assertExcelEqual()
