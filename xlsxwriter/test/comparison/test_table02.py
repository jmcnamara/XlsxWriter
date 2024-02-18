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
        self.set_filename("table02.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with tables."""

        workbook = Workbook(self.got_filename)

        worksheet1 = workbook.add_worksheet()
        worksheet2 = workbook.add_worksheet()

        worksheet1.set_column("B:J", 10.288)
        worksheet2.set_column("C:L", 10.288)

        worksheet1.add_table("B3:E11")
        worksheet1.add_table("G10:J16")
        worksheet1.add_table("C18:F25")

        worksheet2.add_table("I4:L11")
        worksheet2.add_table("C16:H23")

        workbook.close()

        self.assertExcelEqual()
