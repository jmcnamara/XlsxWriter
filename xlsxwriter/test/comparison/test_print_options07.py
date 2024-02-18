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
        self.set_filename("print_options07.xlsx")
        self.ignore_elements = {
            "xl/worksheets/sheet1.xml": ["<pageMargins", "<pageSetup"]
        }

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with print options."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.write(0, 0, "Foo")

        worksheet.set_paper(9)
        worksheet.vertical_dpi = 200

        worksheet.print_black_and_white()

        workbook.close()

        self.assertExcelEqual()
