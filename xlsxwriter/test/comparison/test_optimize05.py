###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2023, John McNamara, jmcnamara@cpan.org
#

from ...workbook import Workbook
from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("optimize05.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(
            self.got_filename, {"constant_memory": True, "in_memory": False}
        )
        worksheet = workbook.add_worksheet()

        bold = workbook.add_format({"bold": 1})
        italic = workbook.add_format({"italic": 1})

        worksheet.write("A1", "Foo", bold)
        worksheet.write("A2", "Bar", italic)
        worksheet.write_rich_string("A3", "a", bold, "bc", "defg")
        worksheet.write_rich_string("B4", "abc", italic, "de", "fg")
        worksheet.write_rich_string("C5", "a", bold, "bc", "defg")
        worksheet.write_rich_string("D6", "abc", italic, "de", "fg")
        worksheet.write_rich_string("E7", "a", bold, "bcdef", "g")
        worksheet.write_rich_string("F8", italic, "abcd", "efg")

        workbook.close()

        self.assertExcelEqual()
