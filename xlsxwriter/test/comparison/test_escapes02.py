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
        self.set_filename("escapes02.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with comments.Check encoding of comments."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.write("A1", "\"<>'&")
        worksheet.write_comment("B2", """<>&"'""")

        worksheet.set_comments_author("""I am '"<>&""")

        workbook.close()

        self.assertExcelEqual()
