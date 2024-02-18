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
        self.set_filename("chart_display_units11.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({"type": "column"})

        chart.axis_ids = [69559424, 69560960]

        data = [
            [10000000, 20000000, 30000000, 20000000, 10000000],
        ]

        worksheet.write_column(0, 0, data[0])

        chart.add_series({"values": "=Sheet1!$A$1:$A$5"})

        chart.set_y_axis({"display_units": "hundreds"})

        worksheet.insert_chart("E9", chart)

        workbook.close()

        self.assertExcelEqual()
