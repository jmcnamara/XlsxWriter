###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2023, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparison_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):

        self.set_filename("quote_name09.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        sheet_name = "Sheet_1"

        worksheet = workbook.add_worksheet(sheet_name)
        chart = workbook.add_chart({"type": "column"})

        chart.axis_ids = [54437760, 59195776]

        data = [
            [1, 2, 3, 4, 5],
            [2, 4, 6, 8, 10],
            [3, 6, 9, 12, 15],
        ]

        worksheet.write_column("A1", data[0])
        worksheet.write_column("B1", data[1])
        worksheet.write_column("C1", data[2])

        worksheet.repeat_rows(0, 1)
        worksheet.set_portrait()
        worksheet.vertical_dpi = 200

        chart.add_series({"values": [sheet_name, 0, 0, 4, 0]})
        chart.add_series({"values": [sheet_name, 0, 1, 4, 1]})
        chart.add_series({"values": [sheet_name, 0, 2, 4, 2]})

        worksheet.insert_chart("E9", chart)

        workbook.close()

        self.assertExcelEqual()
