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
        self.set_filename("types05.xlsx")

        self.ignore_files = [
            "xl/calcChain.xml",
            "[Content_Types].xml",
            "xl/_rels/workbook.xml.rels",
        ]

    def test_write_formula_default(self):
        """Test writing formulas with strings_to_formulas on."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        worksheet.write(0, 0, "=1+1", None, 2)
        worksheet.write_string(1, 0, "=1+1")

        workbook.close()

        self.assertExcelEqual()

    def test_write_formula_implicit(self):
        """Test writing formulas with strings_to_formulas on."""

        workbook = Workbook(self.got_filename, {"strings_to_formulas": True})
        worksheet = workbook.add_worksheet()

        worksheet.write(0, 0, "=1+1", None, 2)
        worksheet.write_string(1, 0, "=1+1")

        workbook.close()

        self.assertExcelEqual()

    def test_write_formula_explicit(self):
        """Test writing formulas with strings_to_formulas off."""

        workbook = Workbook(self.got_filename, {"strings_to_formulas": False})
        worksheet = workbook.add_worksheet()

        worksheet.write_formula(0, 0, "=1+1", None, 2)
        worksheet.write(1, 0, "=1+1")

        workbook.close()

        self.assertExcelEqual()
