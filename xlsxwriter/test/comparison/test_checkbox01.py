###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

from ...workbook import Workbook
from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("checkbox01.xlsx")

    def test_create_file_with_insert_checkbox(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        worksheet.insert_checkbox("E9", False)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_with_insert_checkbox_and_manual_format(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        cell_format = workbook.add_format({"checkbox": True})

        worksheet.insert_checkbox("E9", False, cell_format)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_with_boolean_and_format(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        cell_format = workbook.add_format({"checkbox": True})

        worksheet.write("E9", False, cell_format)

        workbook.close()

        self.assertExcelEqual()
