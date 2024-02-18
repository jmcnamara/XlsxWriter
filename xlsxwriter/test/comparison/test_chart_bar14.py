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
        self.set_filename("chart_bar14.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet1 = workbook.add_worksheet()
        worksheet2 = workbook.add_worksheet()
        worksheet3 = workbook.add_worksheet()
        chartsheet1 = workbook.add_chartsheet()
        chart1 = workbook.add_chart({"type": "bar"})
        chart2 = workbook.add_chart({"type": "bar"})
        chart3 = workbook.add_chart({"type": "column"})

        chart1.axis_ids = [40294272, 40295808]
        chart2.axis_ids = [40261504, 65749760]
        chart3.axis_ids = [65465728, 66388352]

        data = [
            [1, 2, 3, 4, 5],
            [2, 4, 6, 8, 10],
            [3, 6, 9, 12, 15],
        ]

        worksheet2.default_url_format = None

        worksheet2.write_column("A1", data[0])
        worksheet2.write_column("B1", data[1])
        worksheet2.write_column("C1", data[2])

        worksheet2.write("A6", "http://www.perl.com/")

        chart3.add_series({"values": "=Sheet2!$A$1:$A$5"})
        chart3.add_series({"values": "=Sheet2!$B$1:$B$5"})
        chart3.add_series({"values": "=Sheet2!$C$1:$C$5"})

        chart1.add_series({"values": "=Sheet2!$A$1:$A$5"})
        chart1.add_series({"values": "=Sheet2!$B$1:$B$5"})
        chart1.add_series({"values": "=Sheet2!$C$1:$C$5"})

        chart2.add_series({"values": "=Sheet2!$A$1:$A$5"})

        worksheet2.insert_chart("E9", chart1)
        worksheet2.insert_chart("F25", chart2)

        chartsheet1.set_chart(chart3)

        workbook.close()

        self.assertExcelEqual()
