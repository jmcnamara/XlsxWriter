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
        self.set_filename("escapes03.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file.Check encoding of rich strings."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        bold = workbook.add_format({"bold": 1})
        italic = workbook.add_format({"italic": 1})

        worksheet.write("A1", "Foo", bold)
        worksheet.write("A2", "Bar", italic)
        worksheet.write_rich_string("A3", "a", bold, """b"<>'c""", "defg")

        workbook.close()

        self.assertExcelEqual()
