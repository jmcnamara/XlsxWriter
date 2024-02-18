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
        self.set_filename("chart_bar15.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet1 = workbook.add_worksheet()
        chartsheet1 = workbook.add_chartsheet()
        worksheet2 = workbook.add_worksheet()
        chartsheet2 = workbook.add_chartsheet()
        chart1 = workbook.add_chart({"type": "bar"})
        chart2 = workbook.add_chart({"type": "column"})

        chart1.axis_ids = [62576896, 62582784]
        chart2.axis_ids = [65979904, 65981440]

        data = [
            [1, 2, 3, 4, 5],
            [2, 4, 6, 8, 10],
            [3, 6, 9, 12, 15],
        ]

        worksheet1.write_column("A1", data[0])
        worksheet1.write_column("B1", data[1])
        worksheet1.write_column("C1", data[2])

        worksheet2.write_column("A1", data[0])
        worksheet2.write_column("B1", data[1])
        worksheet2.write_column("C1", data[2])

        chart1.add_series({"values": "=Sheet1!$A$1:$A$5"})
        chart1.add_series({"values": "=Sheet1!$B$1:$B$5"})
        chart1.add_series({"values": "=Sheet1!$C$1:$C$5"})

        chart2.add_series({"values": "=Sheet2!$A$1:$A$5"})

        chartsheet1.set_chart(chart1)
        chartsheet2.set_chart(chart2)

        workbook.close()

        self.assertExcelEqual()
