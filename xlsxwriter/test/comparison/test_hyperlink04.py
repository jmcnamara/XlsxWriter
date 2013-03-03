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

        filename = 'hyperlink04.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with hyperlinks."""
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename)

        worksheet1 = workbook.add_worksheet()
        worksheet2 = workbook.add_worksheet()
        worksheet3 = workbook.add_worksheet('Data Sheet')

        worksheet1.write_url('A1', "internal:Sheet2!A1")
        worksheet1.write_url('A3', "internal:Sheet2!A1:A5")
        worksheet1.write_url('A5', "internal:'Data Sheet'!D5", None, 'Some text')
        worksheet1.write_url('E12', "internal:Sheet1!J1")
        worksheet1.write_url('G17', "internal:Sheet2!A1", None, 'Some text')
        worksheet1.write_url('A18', "internal:Sheet2!A1", None, None, 'Tool Tip 1')
        worksheet1.write_url('A20', "internal:Sheet2!A1", None, 'More text', 'Tool Tip 2')

        workbook.close()

        ####################################################

        got, exp = _compare_xlsx_files(self.got_filename,
                                       self.exp_filename,
                                       self.ignore_files,
                                       self.ignore_elements)

        self.assertEqual(got, exp)

    def test_create_file_write(self):
        """Test the creation of a simple XlsxWriter file with hyperlinks with write()"""
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename)

        worksheet1 = workbook.add_worksheet()
        worksheet2 = workbook.add_worksheet()
        worksheet3 = workbook.add_worksheet('Data Sheet')

        worksheet1.write('A1', "internal:Sheet2!A1")
        worksheet1.write('A3', "internal:Sheet2!A1:A5")
        worksheet1.write('A5', "internal:'Data Sheet'!D5", None, 'Some text')
        worksheet1.write('E12', "internal:Sheet1!J1")
        worksheet1.write('G17', "internal:Sheet2!A1", None, 'Some text')
        worksheet1.write('A18', "internal:Sheet2!A1", None, None, 'Tool Tip 1')
        worksheet1.write('A20', "internal:Sheet2!A1", None, 'More text', 'Tool Tip 2')

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
