###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2014, John McNamara, jmcnamara@cpan.org
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

        filename = 'array_formula01.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = ['xl/calcChain.xml',
                             '[Content_Types].xml',
                             'xl/_rels/workbook.xml.rels']
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of an XlsxWriter file with an array formula."""
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename)

        worksheet = workbook.add_worksheet()

        worksheet.write('B1', 0)
        worksheet.write('B2', 0)
        worksheet.write('B3', 0)
        worksheet.write('C1', 0)
        worksheet.write('C2', 0)
        worksheet.write('C3', 0)

        worksheet.write_array_formula(0, 0, 2, 0, '{=SUM(B1:C1*B2:C2)}', None, 0)

        workbook.close()

        ####################################################

        got, exp = _compare_xlsx_files(self.got_filename,
                                       self.exp_filename,
                                       self.ignore_files,
                                       self.ignore_elements)

        self.assertEqual(got, exp)

    def test_create_file_A1(self):
        """
        Test the creation of an XlsxWriter file with an array formula
        and A1 Notation.

        """
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename)

        worksheet = workbook.add_worksheet()

        worksheet.write('B1', 0)
        worksheet.write('B2', 0)
        worksheet.write('B3', 0)
        worksheet.write('C1', 0)
        worksheet.write('C2', 0)
        worksheet.write('C3', 0)

        worksheet.write_array_formula('A1:A3', '{=SUM(B1:C1*B2:C2)}', None, 0)

        workbook.close()

        ####################################################

        got, exp = _compare_xlsx_files(self.got_filename,
                                       self.exp_filename,
                                       self.ignore_files,
                                       self.ignore_elements)

        self.assertEqual(got, exp)

    def test_create_file_kwargs(self):
        """
        Test the creation of an XlsxWriter file with an array formula
        and keyword args
        """
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename)

        worksheet = workbook.add_worksheet()

        worksheet.write('B1', 0)
        worksheet.write('B2', 0)
        worksheet.write('B3', 0)
        worksheet.write('C1', 0)
        worksheet.write('C2', 0)
        worksheet.write('C3', 0)

        worksheet.write_array_formula(first_row=0, first_col=0,
                                      last_row=2, last_col=0,
                                      formula='{=SUM(B1:C1*B2:C2)}')

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
