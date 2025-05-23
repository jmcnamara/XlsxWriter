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
        self.set_filename("array_formula04.xlsx")

        self.ignore_files = [
            "xl/calcChain.xml",
            "[Content_Types].xml",
            "xl/_rels/workbook.xml.rels",
        ]

    def test_create_file(self):
        """Test the creation of an XlsxWriter file with an array formula."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.write_array_formula("A1:A3", "{=SUM(B1:C1*B2:C2)}", None, 0)

        workbook.close()

        self.assertExcelEqual()
