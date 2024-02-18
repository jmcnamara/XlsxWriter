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
        self.set_filename("set_start_page02.xlsx")

        self.ignore_elements = {"xl/worksheets/sheet1.xml": ["<pageMargins"]}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with printer settings."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_start_page(2)
        worksheet.set_paper(9)

        worksheet.vertical_dpi = 200

        worksheet.write("A1", "Foo")

        workbook.close()

        self.assertExcelEqual()
