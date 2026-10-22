###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

from xlsxwriter.workbook import Workbook

from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("firstsheet02.xlsx")

    def test_create_file(self):
        """Test setting the first visible tab in a group of worksheets."""

        workbook = Workbook(self.got_filename)

        for _ in range(20):
            workbook.add_worksheet()

        workbook.worksheets()[1].set_first_sheet()
        workbook.worksheets()[19].activate()

        workbook.close()

        self.assertExcelEqual()
