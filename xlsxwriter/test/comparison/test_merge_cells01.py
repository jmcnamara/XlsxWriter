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
        self.set_filename("merge_cells01.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        cell_format = workbook.add_format({"align": "center"})

        worksheet.set_selection("A4")

        worksheet.merge_range("A1:A2", "col1", cell_format)
        worksheet.merge_range("B1:B2", "col2", cell_format)
        worksheet.merge_range("C1:C2", "col3", cell_format)
        worksheet.merge_range("D1:D2", "col4", cell_format)

        workbook.close()

        self.assertExcelEqual()
