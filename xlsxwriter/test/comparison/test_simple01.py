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

        filename = 'simple01.xlsx'

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
        worksheet = workbook.add_worksheet()

        worksheet.write_string(0, 0, 'Hello')
        worksheet.write_number(1, 0, 123)

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
        worksheet = workbook.add_worksheet()

        worksheet.write('A1', 'Hello')
        worksheet.write('A2', 123)

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
