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
        self.set_filename("custom_ui01.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with custom ribbons."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        workbook.add_custom_ui(self.vba_dir + "customUI-01.xml")
        workbook.add_custom_ui(self.vba_dir + "customUI14-01.xml")

        worksheet.write("A1", "Test")

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_in_memory(self):
        """Test the creation of a simple XlsxWriter file with custom ribbons."""

        workbook = Workbook(self.got_filename, {"in_memory": True})

        worksheet = workbook.add_worksheet()

        workbook.add_custom_ui(self.vba_dir + "customUI-01.xml")
        workbook.add_custom_ui(self.vba_dir + "customUI14-01.xml")

        worksheet.write("A1", "Test")

        workbook.close()

        self.assertExcelEqual()
