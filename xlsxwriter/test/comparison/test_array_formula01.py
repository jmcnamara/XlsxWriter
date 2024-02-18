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
        self.set_filename("array_formula01.xlsx")

        self.ignore_files = [
            "xl/calcChain.xml",
            "[Content_Types].xml",
            "xl/_rels/workbook.xml.rels",
        ]

    def test_create_file(self):
        """Test the creation of an XlsxWriter file with an array formula."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.write("B1", 0)
        worksheet.write("B2", 0)
        worksheet.write("B3", 0)
        worksheet.write("C1", 0)
        worksheet.write("C2", 0)
        worksheet.write("C3", 0)

        worksheet.write_array_formula(0, 0, 2, 0, "{=SUM(B1:C1*B2:C2)}", None, 0)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_A1(self):
        """
        Test the creation of an XlsxWriter file with an array formula
        and A1 Notation.

        """

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.write("B1", 0)
        worksheet.write("B2", 0)
        worksheet.write("B3", 0)
        worksheet.write("C1", 0)
        worksheet.write("C2", 0)
        worksheet.write("C3", 0)

        worksheet.write_array_formula("A1:A3", "{=SUM(B1:C1*B2:C2)}", None, 0)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_kwargs(self):
        """
        Test the creation of an XlsxWriter file with an array formula
        and keyword args
        """

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.write("B1", 0)
        worksheet.write("B2", 0)
        worksheet.write("B3", 0)
        worksheet.write("C1", 0)
        worksheet.write("C2", 0)
        worksheet.write("C3", 0)

        worksheet.write_array_formula(
            first_row=0,
            first_col=0,
            last_row=2,
            last_col=0,
            formula="{=SUM(B1:C1*B2:C2)}",
        )

        workbook.close()

        self.assertExcelEqual()
