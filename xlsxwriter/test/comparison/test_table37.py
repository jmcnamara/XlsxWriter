###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

from xlsxwriter.workbook import Workbook

from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):

        self.set_filename("table37.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with tables."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({"type": "column"})

        worksheet.write(1, 0, 1)
        worksheet.write(2, 0, 2)
        worksheet.write(3, 0, 3)
        worksheet.write(4, 0, 4)
        worksheet.write(5, 0, 5)

        worksheet.write(1, 1, 10)
        worksheet.write(2, 1, 15)
        worksheet.write(3, 1, 20)
        worksheet.write(4, 1, 10)
        worksheet.write(5, 1, 15)

        worksheet.set_column("A:B", 10.288)
        worksheet.add_table("A1:B6")

        chart.axis_ids = [88157568, 89138304]

        chart.add_series(
            {
                "name": "=Sheet1!$B$1",
                "categories": "=Sheet1!$A$2:$A$6",
                "values": "=Sheet1!$B$2:$B$6",
            }
        )

        chart.set_title({"none": True})

        worksheet.insert_chart("E9", chart)

        workbook.close()

        self.assertExcelEqual()
