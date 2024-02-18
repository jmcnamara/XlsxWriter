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
        self.set_filename("escapes05.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file. Check encoding of url strings."""

        workbook = Workbook(self.got_filename)

        # Turn off default URL format for testing.
        workbook.default_url_format = None

        worksheet1 = workbook.add_worksheet("Start")
        worksheet2 = workbook.add_worksheet("A & B")

        worksheet1.write_url("A1", "internal:'A & B'!A1", None, "Jump to A & B")

        workbook.close()

        self.assertExcelEqual()
