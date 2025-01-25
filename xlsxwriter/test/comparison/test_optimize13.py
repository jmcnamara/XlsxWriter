###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

from ...workbook import Workbook
from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("optimize13.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with comments."""

        workbook = Workbook(
            self.got_filename, {"constant_memory": True, "in_memory": False}
        )

        worksheet = workbook.add_worksheet()
        worksheet.write("A1", "Foo")
        worksheet.write_comment("B2", "Some text")

        worksheet.set_comments_author("John")

        workbook.close()

        self.assertExcelEqual()
