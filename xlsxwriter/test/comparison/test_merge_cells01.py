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
        self.set_filename("merge_cells01.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        format = workbook.add_format({"align": "center"})

        worksheet.set_selection("A4")

        worksheet.merge_range("A1:A2", "col1", format)
        worksheet.merge_range("B1:B2", "col2", format)
        worksheet.merge_range("C1:C2", "col3", format)
        worksheet.merge_range("D1:D2", "col4", format)

        workbook.close()

        self.assertExcelEqual()
