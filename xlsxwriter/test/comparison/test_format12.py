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
        self.set_filename("format12.xlsx")

    def test_create_file(self):
        """Test a vertical and horizontal centered format."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        top_left_bottom = workbook.add_format(
            {
                "left": 1,
                "top": 1,
                "bottom": 1,
            }
        )

        top_bottom = workbook.add_format(
            {
                "top": 1,
                "bottom": 1,
            }
        )

        top_left = workbook.add_format(
            {
                "left": 1,
                "top": 1,
            }
        )

        worksheet.write("B2", "test", top_left_bottom)
        worksheet.write("D2", "test", top_left)
        worksheet.write("F2", "test", top_bottom)

        workbook.close()

        self.assertExcelEqual()
