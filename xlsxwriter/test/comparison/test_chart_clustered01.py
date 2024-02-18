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
        self.set_filename("chart_clustered01.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({"type": "column"})

        chart.axis_ids = [45886080, 45928832]

        data = [
            ["Types", "Sub Type", "Value 1", "Value 2", "Value 3"],
            ["Type 1", "Sub Type A", 5000, 8000, 6000],
            ["", "Sub Type B", 2000, 3000, 4000],
            ["", "Sub Type C", 250, 1000, 2000],
            ["Type 2", "Sub Type D", 6000, 6000, 6500],
            ["", "Sub Type E", 500, 300, 200],
        ]

        cat_data = [
            ["Type 1", None, None, "Type 2", None],
            ["Sub Type A", "Sub Type B", "Sub Type C", "Sub Type D", "Sub Type E"],
        ]

        for row_num, row_data in enumerate(data):
            worksheet.write_row(row_num, 0, row_data)

        chart.add_series(
            {
                "name": "=Sheet1!$C$1",
                "categories": "=Sheet1!$A$2:$B$6",
                "values": "=Sheet1!$C$2:$C$6",
                "categories_data": cat_data,
            }
        )

        chart.add_series(
            {
                "name": "=Sheet1!$D$1",
                "categories": "=Sheet1!$A$2:$B$6",
                "values": "=Sheet1!$D$2:$D$6",
            }
        )

        chart.add_series(
            {
                "name": "=Sheet1!$E$1",
                "categories": "=Sheet1!$A$2:$B$6",
                "values": "=Sheet1!$E$2:$E$6",
            }
        )

        worksheet.insert_chart("E9", chart)

        workbook.close()

        self.assertExcelEqual()
