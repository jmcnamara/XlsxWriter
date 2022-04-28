###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2022, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparison_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):

        self.set_filename('autofilter01.xlsx')
        self.set_text_file('autofilter_data.txt')

    def test_create_file(self):
        """
        Test the creation of a simple XlsxWriter file with an autofilter.
        This test corresponds to the following examples/autofilter.py example:
        Example 1. Autofilter without conditions.
        """

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        # Set the autofilter.
        worksheet.autofilter('A1:D51')

        # Open a text file with autofilter example data.
        textfile = open(self.txt_filename)

        # Start writing data from the first worksheet row.
        row = 0

        # Read the text file and write it to the worksheet.
        for line in textfile:

            # Split the input data based on whitespace.
            data = line.strip("\n").split()

            # Convert the number data from the text file.
            for i, item in enumerate(data):
                try:
                    data[i] = float(item)
                except ValueError:
                    pass

            # Write out the row data.
            worksheet.write_row(row, 0, data)

            # Move on to the next worksheet row.
            row += 1

        textfile.close()
        workbook.close()

        self.assertExcelEqual()
