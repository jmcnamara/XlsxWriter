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
        self.set_filename("dynamic_array02.xlsx")

    def test_dynamic_array02_1(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.write_dynamic_array_formula("B1:B1", "=_xlfn.UNIQUE(A1)", None, 0)
        worksheet.write("A1", 0)

        workbook.close()

        self.assertExcelEqual()

    def test_dynamic_array02_2(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.write_dynamic_array_formula("B1", "=_xlfn.UNIQUE(A1)", None, 0)
        worksheet.write("A1", 0)

        workbook.close()

        self.assertExcelEqual()

    def test_dynamic_array02_3(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.write("B1", "=_xlfn.UNIQUE(A1)", None, 0)
        worksheet.write("A1", 0)

        workbook.close()

        self.assertExcelEqual()

    def test_dynamic_array02_4(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.write_dynamic_array_formula("B1", "=UNIQUE(A1)", None, 0)
        worksheet.write("A1", 0)

        workbook.close()

        self.assertExcelEqual()

    def test_dynamic_array02_5(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.write_formula("B1", "=UNIQUE(A1)", None, 0)
        worksheet.write("A1", 0)

        workbook.close()

        self.assertExcelEqual()

    def test_dynamic_array02_6(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.write("B1", "=UNIQUE(A1)", None, 0)
        worksheet.write("A1", 0)

        workbook.close()

        self.assertExcelEqual()

    def test_dynamic_array02_7(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.write_array_formula("B1", "=UNIQUE(A1)", None, 0)
        worksheet.write("A1", 0)

        workbook.close()

        self.assertExcelEqual()
