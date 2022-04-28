###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2022, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparison_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):

        self.set_filename('firstsheet01.xlsx')

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet1 = workbook.add_worksheet()
        worksheet2 = workbook.add_worksheet()
        worksheet3 = workbook.add_worksheet()
        worksheet4 = workbook.add_worksheet()
        worksheet5 = workbook.add_worksheet()
        worksheet6 = workbook.add_worksheet()
        worksheet7 = workbook.add_worksheet()
        worksheet8 = workbook.add_worksheet()
        worksheet9 = workbook.add_worksheet()
        worksheet10 = workbook.add_worksheet()
        worksheet11 = workbook.add_worksheet()
        worksheet12 = workbook.add_worksheet()
        worksheet13 = workbook.add_worksheet()
        worksheet14 = workbook.add_worksheet()
        worksheet15 = workbook.add_worksheet()
        worksheet16 = workbook.add_worksheet()
        worksheet17 = workbook.add_worksheet()
        worksheet18 = workbook.add_worksheet()
        worksheet19 = workbook.add_worksheet()
        worksheet20 = workbook.add_worksheet()

        worksheet8.set_first_sheet()
        worksheet20.activate()

        workbook.close()

        self.assertExcelEqual()
