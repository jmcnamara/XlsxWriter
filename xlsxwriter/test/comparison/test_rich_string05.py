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
        self.set_filename("rich_string05.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_column("A:A", 30)

        bold = workbook.add_format({"bold": 1})
        italic = workbook.add_format({"italic": 1})

        worksheet.write("A1", "Foo", bold)
        worksheet.write("A2", "Bar", italic)
        worksheet.write_rich_string(
            "A3", "This is ", bold, "bold", " and this is ", italic, "italic"
        )

        workbook.close()

        self.assertExcelEqual()
