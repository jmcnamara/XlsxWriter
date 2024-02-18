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
        self.set_filename("table25.xlsx")

    def test_create_file_style_is_none(self):
        """Test the creation of a simple XlsxWriter file with tables."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_column("C:F", 10.288)

        worksheet.add_table("C3:F13", {"style": None})

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_style_is_blank(self):
        """Test the creation of a simple XlsxWriter file with tables."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_column("C:F", 10.288)

        worksheet.add_table("C3:F13", {"style": ""})

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_style_is_none_str(self):
        """Test the creation of a simple XlsxWriter file with tables."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_column("C:F", 10.288)

        worksheet.add_table("C3:F13", {"style": "None"})

        workbook.close()

        self.assertExcelEqual()
