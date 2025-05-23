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
        self.set_filename("unicode_polish_utf8.xlsx")
        self.set_text_file("unicode_polish_utf8.txt")

    def test_create_file(self):
        """Test example file converting Unicode text."""

        # Open the input file with the correct encoding.
        textfile = open(self.txt_filename, mode="r", encoding="utf-8")

        # Create an new Excel file and convert the text data.
        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        # Widen the first column to make the text clearer.
        worksheet.set_column("A:A", 50)

        # Start from the first cell.
        row = 0
        col = 0

        # Read the text file and write it to the worksheet.
        for line in textfile:
            # Ignore the comments in the sample file.
            if line.startswith("#"):
                continue

            # Write any other lines to the worksheet.
            worksheet.write(row, col, line.rstrip("\n"))
            row += 1

        workbook.close()
        textfile.close()

        self.assertExcelEqual()
