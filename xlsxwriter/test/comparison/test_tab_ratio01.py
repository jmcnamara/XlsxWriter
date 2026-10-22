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
        self.set_filename("tab_ratio01.xlsx")

    def test_create_file(self):
        """Test the default tab ratio."""

        workbook = Workbook(self.got_filename)

        workbook.add_worksheet()

        workbook.close()

        self.assertExcelEqual()
