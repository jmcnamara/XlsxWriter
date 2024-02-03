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
        self.set_filename("format24.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with automatic color."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        format1 = workbook.add_format(
            {"rotation": 270, "indent": 1, "align": "center", "valign": "top"}
        )

        worksheet.set_row(0, 75)

        worksheet.write(0, 0, "ABCD", format1)

        workbook.close()

        self.assertExcelEqual()
