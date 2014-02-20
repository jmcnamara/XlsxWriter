###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2014, John McNamara, jmcnamara@cpan.org
#

import unittest
import os
from datetime import datetime
from ...workbook import Workbook
from ..helperfunctions import _compare_xlsx_files


class TestCompareXLSXFiles(unittest.TestCase):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.maxDiff = None

        filename = 'default_format02.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file_user_date_format(self):
        """Test write_datetime with explicit date format."""
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_column(0, 0, 12)

        format1 = workbook.add_format({'num_format': 'dd\\ mm\\ yy'})

        date1 = datetime.strptime('2013-07-25', "%Y-%m-%d")

        worksheet.write_datetime(0, 0, date1, format1)

        workbook.close()

        ####################################################

        got, exp = _compare_xlsx_files(self.got_filename,
                                       self.exp_filename,
                                       self.ignore_files,
                                       self.ignore_elements)

        self.assertEqual(got, exp)

    def test_create_file_default_date_format(self):
        """Test write_datetime with default date format."""
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename, {'default_date_format': 'dd\\ mm\\ yy'})

        worksheet = workbook.add_worksheet()

        worksheet.set_column(0, 0, 12)

        date1 = datetime.strptime('2013-07-25', "%Y-%m-%d")

        worksheet.write_datetime(0, 0, date1)

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
