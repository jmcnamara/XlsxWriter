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
        self.set_filename("print_across01.xlsx")

        self.ignore_files = [
            "xl/printerSettings/printerSettings1.bin",
            "xl/worksheets/_rels/sheet1.xml.rels",
        ]
        self.ignore_elements = {
            "[Content_Types].xml": ['<Default Extension="bin"'],
            "xl/worksheets/sheet1.xml": ["<pageMargins"],
        }

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with fit across."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.print_across()
        worksheet.set_paper(9)

        worksheet.write("A1", "Foo")

        workbook.close()

        self.assertExcelEqual()
