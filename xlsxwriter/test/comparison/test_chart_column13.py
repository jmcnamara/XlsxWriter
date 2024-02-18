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
        self.set_filename("chart_column13.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({"type": "column"})

        chart.axis_ids = [60474496, 78612736]

        worksheet.write("A1", "1.1_1")
        worksheet.write("B1", "2.2_2")
        worksheet.write("A2", 1)
        worksheet.write("B2", 2)

        chart.add_series(
            {"categories": "=Sheet1!$A$1:$B$1", "values": "=Sheet1!$A$2:$B$2"}
        )

        worksheet.insert_chart("E9", chart)

        workbook.close()

        self.assertExcelEqual()
