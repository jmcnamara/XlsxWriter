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

        filename = 'simple07.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_write_nan(self):
        """Test write with NAN/INF. Issue #30"""
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename)
        worksheet = workbook.add_worksheet()

        worksheet.write_string(0, 0, 'Foo')
        worksheet.write_number(1, 0, 123)
        worksheet.write_string(2, 0, 'NAN')
        worksheet.write_string(3, 0, 'nan')
        worksheet.write_string(4, 0, 'INF')
        worksheet.write_string(5, 0, 'infinity')

        workbook.close()

        ####################################################

        got, exp = _compare_xlsx_files(self.got_filename,
                                       self.exp_filename,
                                       self.ignore_files,
                                       self.ignore_elements)

        self.assertEqual(got, exp)

    def test_create_file_in_memory(self):
        """Test write with NAN/INF. Issue #30"""
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename, {'in_memory': True})
        worksheet = workbook.add_worksheet()

        worksheet.write_string(0, 0, 'Foo')
        worksheet.write_number(1, 0, 123)
        worksheet.write_string(2, 0, 'NAN')
        worksheet.write_string(3, 0, 'nan')
        worksheet.write_string(4, 0, 'INF')
        worksheet.write_string(5, 0, 'infinity')

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
