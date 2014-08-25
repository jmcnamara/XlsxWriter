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

        filename = 'optimize11.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file_no_close(self):
        """Test the creation of a simple XlsxWriter file."""
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename, {'constant_memory': True, 'in_memory': False})

        for i in range(1, 10):
            worksheet = workbook.add_worksheet()
            worksheet.write('A1', 'Hello 1')
            worksheet.write('A2', 'Hello 2')
            worksheet.write('A4', 'Hello 3')

        workbook.close()

        ####################################################

        got, exp = _compare_xlsx_files(self.got_filename,
                                       self.exp_filename,
                                       self.ignore_files,
                                       self.ignore_elements)

        self.assertEqual(got, exp)

    def test_create_file_with_close(self):
        """Test the creation of a simple XlsxWriter file."""
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename, {'constant_memory': True, 'in_memory': False})

        for i in range(1, 10):
            worksheet = workbook.add_worksheet()
            worksheet.write('A1', 'Hello 1')
            worksheet.write('A2', 'Hello 2')
            worksheet.write('A4', 'Hello 3')
            worksheet._opt_close()

        workbook.close()

        ####################################################

        got, exp = _compare_xlsx_files(self.got_filename,
                                       self.exp_filename,
                                       self.ignore_files,
                                       self.ignore_elements)

        self.assertEqual(got, exp)

    def test_create_file_with_reopen(self):
        """Test the creation of a simple XlsxWriter file."""
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename, {'constant_memory': True, 'in_memory': False})

        for i in range(1, 10):
            worksheet = workbook.add_worksheet()
            worksheet.write('A1', 'Hello 1')
            worksheet._opt_close()
            worksheet._opt_reopen()
            worksheet.write('A2', 'Hello 2')
            worksheet._opt_close()
            worksheet._opt_reopen()
            worksheet.write('A4', 'Hello 3')
            worksheet._opt_close()
            worksheet._opt_reopen()
            worksheet._opt_close()

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
