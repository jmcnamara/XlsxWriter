###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2023, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparison_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("page_view01.xlsx")

        self.ignore_files = [
            "xl/printerSettings/printerSettings1.bin",
            "xl/worksheets/_rels/sheet1.xml.rels",
        ]
        self.ignore_elements = {
            "[Content_Types].xml": ['<Default Extension="bin"'],
            "xl/worksheets/sheet1.xml": ["<pageMargins", "<pageSetup"],
        }

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with print options."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_page_view()

        worksheet.write("A1", "Foo")

        workbook.close()

        self.assertExcelEqual()
