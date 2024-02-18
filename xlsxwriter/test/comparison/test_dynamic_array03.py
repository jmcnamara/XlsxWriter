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
        self.set_filename("dynamic_array03.xlsx")

    def test_future_function01(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.write_formula("A1", "=1+_xlfn.XOR(1)", None, 2)

        workbook.close()

        self.assertExcelEqual()

    def test_future_function02(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        worksheet.use_future_functions = True

        worksheet.write_formula("A1", "=1+XOR(1)", None, 2)

        workbook.close()

        self.assertExcelEqual()

    def test_future_function03(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename, {"use_future_functions": True})

        worksheet = workbook.add_worksheet()

        worksheet.write_formula("A1", "=1+XOR(1)", None, 2)

        workbook.close()

        self.assertExcelEqual()
