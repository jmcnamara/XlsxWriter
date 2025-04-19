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
        self.set_filename("checkbox06.xlsx")

    def test_create_file_with_insert_checkbox(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        worksheet.write(0, 0, "Col1")
        worksheet.write(1, 0, 1)
        worksheet.write(2, 0, 2)
        worksheet.write(3, 0, 3)
        worksheet.write(4, 0, 4)

        worksheet.write(0, 1, "Col2")
        worksheet.insert_checkbox(1, 1, True)
        worksheet.insert_checkbox(2, 1, False)
        worksheet.insert_checkbox(3, 1, False)
        worksheet.insert_checkbox(4, 1, True)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_with_boolean_and_format(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        cell_format = workbook.add_format({"checkbox": True})

        worksheet.write(0, 0, "Col1")
        worksheet.write(1, 0, 1)
        worksheet.write(2, 0, 2)
        worksheet.write(3, 0, 3)
        worksheet.write(4, 0, 4)

        worksheet.write(0, 1, "Col2")
        worksheet.write(1, 1, True, cell_format)
        worksheet.write(2, 1, False, cell_format)
        worksheet.write(3, 1, False, cell_format)
        worksheet.write(4, 1, True, cell_format)

        workbook.close()

        self.assertExcelEqual()
