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
        self.set_filename("utf8_11.xlsx")

    def test_create_file(self):
        """Test the creation of an XlsxWriter file with utf-8 strings."""

        workbook = Workbook(self.got_filename, {"strings_to_urls": True})

        worksheet = workbook.add_worksheet()

        worksheet.write("A1", "１２３４５")

        workbook.close()

        self.assertExcelEqual()
