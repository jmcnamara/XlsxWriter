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
        self.set_filename("array_formula03.xlsx")

        self.ignore_files = [
            "xl/calcChain.xml",
            "[Content_Types].xml",
            "xl/_rels/workbook.xml.rels",
        ]

    def test_create_file_write_formula(self):
        """Test the creation of an XlsxWriter file with an array formula."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        data = [0, 0, 0]

        worksheet.write_column("B1", data)
        worksheet.write_column("C1", data)

        worksheet.write_formula("A1", "{=SUM(B1:C1*B2:C2)}", None)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_write(self):
        """Test the creation of an XlsxWriter file with an array formula."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        data = [0, 0, 0]

        worksheet.write_column("B1", data)
        worksheet.write_column("C1", data)

        worksheet.write("A1", "{=SUM(B1:C1*B2:C2)}", None)

        workbook.close()

        self.assertExcelEqual()
