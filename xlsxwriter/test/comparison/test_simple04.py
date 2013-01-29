###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013, John McNamara, jmcnamara@cpan.org
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

        filename = 'simple04.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test dates and times."""
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_column(0, 0, 12)

        format1 = workbook.add_format({'num_format': 20})
        format2 = workbook.add_format({'num_format': 14})

        date1 = datetime.strptime('12:00', "%H:%M")
        date2 = datetime.strptime('2013-01-27', "%Y-%m-%d")

        worksheet.write_datetime(0, 0, date1, format1)
        worksheet.write_datetime(1, 0, date2, format2)

        workbook.close()

        ####################################################

        got, exp = _compare_xlsx_files(self.got_filename,
                                       self.exp_filename,
                                       self.ignore_files,
                                       self.ignore_elements)

        self.assertEqual(got, exp)

    def test_create_file_write(self):
        """Test dates and times with write() method."""
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_column(0, 0, 12)

        format1 = workbook.add_format({'num_format': 20})
        format2 = workbook.add_format({'num_format': 14})

        date1 = datetime.strptime('12:00', "%H:%M")
        date2 = datetime.strptime('2013-01-27', "%Y-%m-%d")

        worksheet.write(0, 0, date1, format1)
        worksheet.write(1, 0, date2, format2)

        workbook.close()

        ####################################################

        got, exp = _compare_xlsx_files(self.got_filename,
                                       self.exp_filename,
                                       self.ignore_files,
                                       self.ignore_elements)

        self.assertEqual(got, exp)

    def test_create_file_A1(self):
        """Test dates and times in A1 notation."""
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_column(0, 0, 12)

        format1 = workbook.add_format({'num_format': 20})
        format2 = workbook.add_format({'num_format': 14})

        date1 = datetime.strptime('12:00', "%H:%M")
        date2 = datetime.strptime('2013-01-27', "%Y-%m-%d")

        worksheet.write_datetime('A1', date1, format1)
        worksheet.write_datetime('A2', date2, format2)

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
