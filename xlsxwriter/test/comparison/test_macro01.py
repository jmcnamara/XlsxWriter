###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparison_test import ExcelComparisonTest
from ...workbook import Workbook
from io import BytesIO


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("macro01.xlsm")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        workbook.add_vba_project(self.vba_dir + "vbaProject01.bin")

        worksheet.write("A1", 123)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_in_memory(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename, {"in_memory": True})

        worksheet = workbook.add_worksheet()

        workbook.add_vba_project(self.vba_dir + "vbaProject01.bin")

        worksheet.write("A1", 123)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_bytes_io(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        vba_file = open(self.vba_dir + "vbaProject01.bin", "rb")
        vba_data = BytesIO(vba_file.read())
        vba_file.close()

        workbook.add_vba_project(vba_data, True)

        worksheet.write("A1", 123)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_bytes_io_in_memory(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename, {"in_memory": True})

        worksheet = workbook.add_worksheet()

        vba_file = open(self.vba_dir + "vbaProject01.bin", "rb")
        vba_data = BytesIO(vba_file.read())
        vba_file.close()

        workbook.add_vba_project(vba_data, True)

        worksheet.write("A1", 123)

        workbook.close()

        self.assertExcelEqual()
