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
        self.set_filename("chart_name03.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        chart1 = workbook.add_chart({"type": "line", "name": "New 1"})

        chart2 = workbook.add_chart({"type": "line", "name": "New 2"})

        chart1.axis_ids = [44271104, 45703168]
        chart2.axis_ids = [80928128, 80934400]

        data = [
            [1, 2, 3, 4, 5],
            [2, 4, 6, 8, 10],
            [3, 6, 9, 12, 15],
        ]

        worksheet.write_column("A1", data[0])
        worksheet.write_column("B1", data[1])
        worksheet.write_column("C1", data[2])

        chart1.add_series({"values": "=Sheet1!$A$1:$A$5"})
        chart1.add_series({"values": "=Sheet1!$B$1:$B$5"})
        chart1.add_series({"values": "=Sheet1!$C$1:$C$5"})

        chart2.add_series({"values": "=Sheet1!$A$1:$A$5"})
        chart2.add_series({"values": "=Sheet1!$B$1:$B$5"})
        chart2.add_series({"values": "=Sheet1!$C$1:$C$5"})

        worksheet.insert_chart("E9", chart1)
        worksheet.insert_chart("E24", chart2)

        workbook.close()

        self.assertExcelEqual()
