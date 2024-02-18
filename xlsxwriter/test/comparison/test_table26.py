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
        self.set_filename("table26.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with tables."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_column("C:D", 10.288)
        worksheet.set_column("F:G", 10.288)

        worksheet.add_table("C2:D3")
        worksheet.add_table("F3:G3", {"header_row": 0})

        # These tables should be ignored since the ranges are incorrect.
        import warnings

        warnings.filterwarnings("ignore")
        worksheet.add_table("I2:J2")
        worksheet.add_table("L3:M3", {"header_row": 1})

        workbook.close()

        self.assertExcelEqual()
