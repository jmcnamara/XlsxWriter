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
        self.set_filename("autofilter00.xlsx")
        self.set_text_file("autofilter_data.txt")

    def test_create_file(self):
        """
        Test the creation of a simple XlsxWriter file with an autofilter.
        This test is the base comparison. It has data but no autofilter.
        """

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        # Open a text file with autofilter example data.
        with open(self.txt_filename) as textfile:
            # Read the text file and write it to the worksheet.
            for row, line in enumerate(textfile):
                # Split the input data based on whitespace.
                data = line.strip("\n").split()

                # Convert the number data from the text file.
                for i, item in enumerate(data):
                    try:
                        data[i] = float(item)
                    except ValueError:
                        pass

                for col in range(len(data)):
                    worksheet.write(row, col, data[col])

        workbook.close()

        self.assertExcelEqual()
