###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

from xlsxwriter.workbook import Workbook

from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("macro05.xlsm")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with custom ribbon tab."""

        workbook = Workbook(self.got_filename)
        
        worksheet = workbook.add_worksheet()
        
        workbook.add_custom_ui(self.vba_dir + "customUI-01.xml", version=2006)
        workbook.add_custom_ui(self.vba_dir + "customUI14-01.xml", version=2007)

        workbook.add_signed_vba_project(
            self.vba_dir + "vbaProject06.bin",
            self.vba_dir + "vbaProject06Signature.bin",
        )

        worksheet.write("A1", "Test")

        workbook.close()

        self.assertExcelEqual()
