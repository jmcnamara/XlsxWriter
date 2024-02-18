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
        self.set_filename("rich_string12.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_column("A:A", 30)
        worksheet.set_row(2, 60)

        bold = workbook.add_format({"bold": 1})
        italic = workbook.add_format({"italic": 1})
        wrap = workbook.add_format({"text_wrap": 1})

        worksheet.write("A1", "Foo", bold)
        worksheet.write("A2", "Bar", italic)

        worksheet.write_rich_string(
            "A3", "This is\n", bold, "bold\n", "and this is\n", italic, "italic", wrap
        )

        workbook.close()

        self.assertExcelEqual()
