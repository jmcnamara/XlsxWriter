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
        self.set_filename("optimize01.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(
            self.got_filename,
            {"constant_memory": True, "strings_to_numbers": True, "in_memory": False},
        )
        worksheet = workbook.add_worksheet()

        worksheet.write("A1", "Hello")
        worksheet.write("A2", "123")

        workbook.close()

        self.assertExcelEqual()
