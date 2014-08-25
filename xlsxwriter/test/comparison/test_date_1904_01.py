###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2014, John McNamara, jmcnamara@cpan.org
#

import unittest
import os
from datetime import date
from ...workbook import Workbook
from ..helperfunctions import _compare_xlsx_files


class TestCompareXLSXFiles(unittest.TestCase):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.maxDiff = None

        filename = 'date_1904_01.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a XlsxWriter file with date times in 1900 and1904 epochs."""
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename)

        worksheet = workbook.add_worksheet()
        format1 = workbook.add_format({'num_format': 14})

        worksheet.set_column('A:A', 12)

        worksheet.write_datetime('A1', date(1900, 1, 1), format1)
        worksheet.write_datetime('A2', date(1902, 9, 26), format1)
        worksheet.write_datetime('A3', date(1913, 9, 8), format1)
        worksheet.write_datetime('A4', date(1927, 5, 18), format1)
        worksheet.write_datetime('A5', date(2173, 10, 14), format1)
        worksheet.write_datetime('A6', date(4637, 11, 26), format1)

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
