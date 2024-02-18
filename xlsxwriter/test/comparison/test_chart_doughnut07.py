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
        self.set_filename("chart_doughnut07.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart_doughnut = workbook.add_chart({"type": "doughnut"})

        worksheet.write_column("H2", ["Donut", 25, 50, 25, 100])
        worksheet.write_column("I2", ["Pie", 75, 1, 124])

        chart_doughnut.add_series(
            {
                "name": "=Sheet1!$H$2",
                "values": "=Sheet1!$H$3:$H$6",
            }
        )

        chart_doughnut.add_series(
            {
                "name": "=Sheet1!$I$2",
                "values": "=Sheet1!$I$3:$I$6",
            }
        )

        worksheet.insert_chart("E9", chart_doughnut)

        workbook.close()

        self.assertExcelEqual()
