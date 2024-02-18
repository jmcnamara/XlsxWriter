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
        self.set_filename("chart_bar02.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet1 = workbook.add_worksheet()
        worksheet2 = workbook.add_worksheet()
        chart = workbook.add_chart({"type": "bar"})

        chart.axis_ids = [93218304, 93219840]

        data = [
            [1, 2, 3, 4, 5],
            [2, 4, 6, 8, 10],
            [3, 6, 9, 12, 15],
        ]

        worksheet1.write("A1", "Foo")

        worksheet2.write_column("A1", data[0])
        worksheet2.write_column("B1", data[1])
        worksheet2.write_column("C1", data[2])

        chart.add_series(
            {
                "categories": "Sheet2!$A$1:$A$5",
                "values": "Sheet2!$B$1:$B$5",
            }
        )

        chart.add_series(
            {
                "categories": "Sheet2!$A$1:$A$5",
                "values": "Sheet2!$C$1:$C$5",
            }
        )
        worksheet2.insert_chart("E9", chart)

        workbook.close()

        self.assertExcelEqual()
