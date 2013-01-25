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


class TestCreateXLSXFile(unittest.TestCase):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.maxDiff = None

        filename = 'simple02.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_members = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple workbook."""
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename)

        worksheet1 = workbook.add_worksheet()
        workbook.add_worksheet('Data Sheet')
        worksheet3 = workbook.add_worksheet()

        bold = workbook.add_format({'bold': 1})

        worksheet1.write_string(0, 0, 'Foo')
        worksheet1.write_number(1, 0, 123)

        worksheet3.write_string(1, 1, 'Foo')
        worksheet3.write_string(2, 1, 'Bar', bold)
        worksheet3.write_number(3, 2, 234)

        workbook.close()

        ####################################################

        exp, got = _compare_xlsx_files(self.got_filename,
                                       self.exp_filename,
                                       self.ignore_members,
                                       self.ignore_elements)

        self.assertEqual(got, exp)

    def test_create_file_A1(self):
        """Test the creation of a simple workbook with A1 notation."""
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename)

        worksheet1 = workbook.add_worksheet()
        workbook.add_worksheet('Data Sheet')
        worksheet3 = workbook.add_worksheet()

        bold = workbook.add_format({'bold': 1})

        worksheet1.write('A1', 'Foo')
        worksheet1.write('A2', 123)

        worksheet3.write('B2', 'Foo')
        worksheet3.write('B3', 'Bar', bold)
        worksheet3.write('C4', 234)

        workbook.close()

        ####################################################

        exp, got = _compare_xlsx_files(self.got_filename,
                                       self.exp_filename,
                                       self.ignore_members,
                                       self.ignore_elements)

        self.assertEqual(got, exp)

    def tearDown(self):
        # Cleanup.
        os.remove(self.got_filename)


if __name__ == '__main__':
    unittest.main()
