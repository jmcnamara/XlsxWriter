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
        self.set_filename("types01.xlsx")

    def test_write_number_as_text(self):
        """Test writing numbers as text."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        worksheet.write_string(0, 0, "Hello")
        worksheet.write_string(1, 0, "123")

        workbook.close()

        self.assertExcelEqual()

    def test_write_number_as_text_with_write(self):
        """Test writing numbers as text using write() without conversion."""

        workbook = Workbook(self.got_filename, {"strings_to_numbers": False})
        worksheet = workbook.add_worksheet()

        worksheet.write(0, 0, "Hello")
        worksheet.write(1, 0, "123")

        workbook.close()

        self.assertExcelEqual()
