###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013, John McNamara, jmcnamara@cpan.org
#

import unittest
import os
from ...workbook import Workbook
from ..helperfunctions import _compare_xlsx_files


class TestCompareXLSXFiles(unittest.TestCase):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.maxDiff = None

        filename = 'tutorial02.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = ['xl/calcChain.xml',
                             '[Content_Types].xml',
                             'xl/_rels/workbook.xml.rels']
        self.ignore_elements = {}

    def test_create_file(self):
        """Example spreadsheet used in the tutorial 2."""
        filename = self.got_filename

        ####################################################

        # Create a workbook and add a worksheet.
        workbook = Workbook(filename)
        worksheet = workbook.add_worksheet()

        # Add a bold format to use to highlight cells.
        bold = workbook.add_format({'bold': True})

        # Add a number format for cells with money.
        money_format = workbook.add_format({'num_format': '\\$#,##0'})

        # Write some data headers.
        worksheet.write('A1', 'Item', bold)
        worksheet.write('B1', 'Cost', bold)

        # Some data we want to write to the worksheet.
        expenses = (
            ['Rent', 1000],
            ['Gas', 100],
            ['Food', 300],
            ['Gym', 50],
        )

        # Start from the first cell below the headers.
        row = 1
        col = 0

        # Iterate over the data and write it out row by row.
        for item, cost in (expenses):
            worksheet.write(row, col, item)
            worksheet.write(row, col + 1, cost, money_format)
            row += 1

        # Write a total using a formula.
        worksheet.write(row, 0, 'Total', bold)
        worksheet.write(row, 1, '=SUM(B2:B5)', money_format, 1450)

        workbook.close()

        ####################################################

        got, exp = _compare_xlsx_files(self.got_filename,
                                       self.exp_filename,
                                       self.ignore_files,
                                       self.ignore_elements)

        self.assertEqual(got, exp)

    def tearDown(self):
        # Cleanup.
        if os.path.exists(self.got_filename):
            os.remove(self.got_filename)


if __name__ == '__main__':
    unittest.main()
