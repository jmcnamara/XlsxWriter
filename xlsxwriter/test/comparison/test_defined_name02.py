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
        self.set_filename("defined_name02.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with defined names."""

        workbook = Workbook(self.got_filename)

        worksheet1 = workbook.add_worksheet("sheet One")

        workbook.define_name("Sales", "='sheet One'!$G$1:$H$10")

        workbook.close()

        self.assertExcelEqual()
