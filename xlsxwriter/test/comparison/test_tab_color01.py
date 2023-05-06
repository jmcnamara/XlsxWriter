###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2023, John McNamara, jmcnamara@cpan.org
#

from ...workbook import Workbook
from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("tab_color01.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with a colored tab."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.write("A1", "Foo")
        worksheet.set_tab_color("red")

        workbook.close()

        self.assertExcelEqual()
