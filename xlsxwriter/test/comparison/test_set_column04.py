###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

from xlsxwriter.workbook import Workbook

from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("set_column04.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        bold = workbook.add_format({"bold": 1})
        italic = workbook.add_format({"italic": 1})
        bold_italic = workbook.add_format({"bold": 1, "italic": 1})

        data = [
            [1, 2, 3, 4, 5],
            [2, 4, 6, 8, 10],
            [3, 6, 9, 12, 15],
        ]

        worksheet.write("A1", "Foo", italic)
        worksheet.write("B1", "Bar", bold)
        worksheet.write_column("A2", data[0])
        worksheet.write_column("B2", data[1])
        worksheet.write_column("C2", data[2])

        worksheet.set_row(12, None, italic)
        worksheet.set_column("F:F", None, bold)

        worksheet.write("F13", None, bold_italic)

        workbook.close()

        self.assertExcelEqual()
