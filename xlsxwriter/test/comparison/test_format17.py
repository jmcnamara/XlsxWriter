###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

from xlsxwriter.color import Color
from xlsxwriter.workbook import Workbook

from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("format17.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with a pattern only."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        pattern = workbook.add_format({"pattern": 2, "fg_color": "red"})

        worksheet.write("A1", "", pattern)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_with_color(self):
        """Test the creation of a simple XlsxWriter file with a pattern only."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        pattern = workbook.add_format({"pattern": 2, "fg_color": Color("red")})

        worksheet.write("A1", "", pattern)

        workbook.close()

        self.assertExcelEqual()
