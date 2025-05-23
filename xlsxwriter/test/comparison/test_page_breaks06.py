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
        self.set_filename("page_breaks06.xlsx")

        self.ignore_files = [
            "xl/printerSettings/printerSettings1.bin",
            "xl/worksheets/_rels/sheet1.xml.rels",
        ]
        self.ignore_elements = {
            "[Content_Types].xml": ['<Default Extension="bin"'],
            "xl/worksheets/sheet1.xml": ["<pageMargins", "<pageSetup"],
        }

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with page breaks."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_h_pagebreaks([1, 5, 13, 8, 8, 8])
        worksheet.set_v_pagebreaks([8, 3, 1, 0])

        worksheet.write("A1", "Foo")

        workbook.close()

        self.assertExcelEqual()
