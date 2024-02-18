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
        self.set_filename("button07.xlsm")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        workbook.set_vba_name()
        worksheet.set_vba_name()

        worksheet.insert_button("C2", {"macro": "say_hello", "caption": "Hello"})

        workbook.add_vba_project(self.vba_dir + "vbaProject02.bin")

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_explicit_vba_names(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        workbook.set_vba_name("ThisWorkbook")
        worksheet.set_vba_name("Sheet1")

        worksheet.insert_button("C2", {"macro": "say_hello", "caption": "Hello"})

        workbook.add_vba_project(self.vba_dir + "vbaProject02.bin")

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_implicit_vba_names(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.insert_button("C2", {"macro": "say_hello", "caption": "Hello"})

        workbook.add_vba_project(self.vba_dir + "vbaProject02.bin")

        workbook.close()

        self.assertExcelEqual()
