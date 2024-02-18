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
        self.set_filename("chartsheet08.xlsx")

        self.ignore_files = ["xl/drawings/drawing1.xml"]

    def test_create_file(self):
        """Test the worksheet properties of an XlsxWriter chartsheet file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chartsheet = workbook.add_chartsheet()

        chart = workbook.add_chart({"type": "bar"})

        chart.axis_ids = [61297792, 61299328]

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

        chartsheet.set_margins(
            left="0.51181102362204722",
            right="0.51181102362204722",
            top="0.55118110236220474",
            bottom="0.94488188976377963",
        )
        chartsheet.set_header("&CPage &P", "0.11811023622047245")
        chartsheet.set_footer("&C&A", "0.11811023622047245")
        chartsheet.set_paper(9)
        chartsheet.set_portrait()

        chartsheet.horizontal_dpi = 200
        chartsheet.vertical_dpi = 200

        chartsheet.set_chart(chart)

        workbook.close()

        self.assertExcelEqual()
