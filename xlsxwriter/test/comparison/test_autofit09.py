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
        self.set_filename("autofit09.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        text_wrap = workbook.add_format({"text_wrap": True})

        worksheet.write_string(0, 0, "Hello\nFoo", text_wrap)
        worksheet.write_string(2, 2, "Foo\nBamboo\nBar", text_wrap)

        worksheet.set_row(0, 33)
        worksheet.set_row(2, 48)

        worksheet.autofit()

        workbook.close()

        self.assertExcelEqual()
