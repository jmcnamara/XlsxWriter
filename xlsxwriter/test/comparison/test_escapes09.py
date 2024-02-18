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
        self.set_filename("escapes09.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({"type": "line"})

        chart.axis_ids = [52721920, 53133312]

        worksheet.write(0, 0, "Data\x1b[32m1")
        worksheet.write(1, 0, "Data\x1b[32m2")
        worksheet.write(2, 0, "Data\x1b[32m3")
        worksheet.write(3, 0, "Data\x1b[32m4")

        worksheet.write(0, 1, 10)
        worksheet.write(1, 1, 20)
        worksheet.write(2, 1, 10)
        worksheet.write(3, 1, 30)

        chart.add_series(
            {"categories": "=Sheet1!$A$1:$A$4", "values": "=Sheet1!$B$1:$B$4"}
        )

        worksheet.insert_chart("E9", chart)

        workbook.close()

        self.assertExcelEqual()
