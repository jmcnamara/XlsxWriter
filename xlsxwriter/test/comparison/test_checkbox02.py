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
        self.set_filename("checkbox02.xlsx")

    def test_create_file_with_insert_checkbox(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        worksheet.insert_checkbox(0, 0, False)
        worksheet.insert_checkbox(2, 2, True)
        worksheet.insert_checkbox(8, 4, False)
        worksheet.insert_checkbox(9, 4, True)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_with_insert_checkbox_and_manual_format(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        cell_format = workbook.add_format({"checkbox": True})

        worksheet.insert_checkbox(0, 0, False, cell_format)
        worksheet.insert_checkbox(2, 2, True, cell_format)
        worksheet.insert_checkbox(8, 4, False, cell_format)
        worksheet.insert_checkbox(9, 4, True, cell_format)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_with_boolean_and_format(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        cell_format = workbook.add_format({"checkbox": True})

        worksheet.write(0, 0, False, cell_format)
        worksheet.write(2, 2, True, cell_format)
        worksheet.write(8, 4, False, cell_format)
        worksheet.write(9, 4, True, cell_format)

        workbook.close()

        self.assertExcelEqual()
