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
        self.set_filename("chart_combined11.xlsx")
        self.ignore_elements = {"xl/charts/chart1.xml": ["<c:dispBlanksAs"]}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart_doughnut = workbook.add_chart({"type": "doughnut"})
        chart_pie = workbook.add_chart({"type": "pie"})

        worksheet.write_column("H2", ["Donut", 25, 50, 25, 100])
        worksheet.write_column("I2", ["Pie", 75, 1, 124])

        chart_doughnut.add_series(
            {
                "name": "=Sheet1!$H$2",
                "values": "=Sheet1!$H$3:$H$6",
                "points": [
                    {"fill": {"color": "#FF0000"}},
                    {"fill": {"color": "#FFC000"}},
                    {"fill": {"color": "#00B050"}},
                    {"fill": {"none": True}},
                ],
            }
        )

        chart_doughnut.set_rotation(270)
        chart_doughnut.set_legend({"none": True})
        chart_doughnut.set_chartarea(
            {
                "border": {"none": True},
                "fill": {"none": True},
            }
        )

        chart_pie.add_series(
            {
                "name": "=Sheet1!$I$2",
                "values": "=Sheet1!$I$3:$I$6",
                "points": [
                    {"fill": {"none": True}},
                    {"fill": {"color": "#FF0000"}},
                    {"fill": {"none": True}},
                ],
            }
        )

        chart_pie.set_rotation(270)

        chart_doughnut.combine(chart_pie)

        worksheet.insert_chart("A1", chart_doughnut)

        workbook.close()

        self.assertExcelEqual()
