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
        self.set_filename("table03.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with tables."""

        workbook = Workbook(self.got_filename)

        # Turn off default URL format for testing.
        workbook.default_url_format = None

        worksheet = workbook.add_worksheet()

        worksheet.set_column("C:F", 10.288)

        worksheet.add_table("C3:F13")

        worksheet.write("A1", "http://perl.com/")

        workbook.close()

        self.assertExcelEqual()
