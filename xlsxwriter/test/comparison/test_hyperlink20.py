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

        filename = 'hyperlink20.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_hyperlink_formating_explicit(self):
        """Test the creation of a simple XlsxWriter file with hyperlinks. This example has link formatting."""
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename)

        # Simulate custom colour for testing.
        workbook.custom_colors = ['FF0000FF']

        worksheet = workbook.add_worksheet()
        format1 = workbook.add_format({'color': 'blue', 'underline': 1})
        format2 = workbook.add_format({'color': 'red', 'underline': 1})

        worksheet.write_url('A1', 'http://www.python.org/1', format1)
        worksheet.write_url('A2', 'http://www.python.org/2', format2)

        workbook.close()

        ####################################################

        got, exp = _compare_xlsx_files(self.got_filename,
                                       self.exp_filename,
                                       self.ignore_files,
                                       self.ignore_elements)

        self.assertEqual(got, exp)

    def test_hyperlink_formating_implicit(self):
        """Test the creation of a simple XlsxWriter file with hyperlinks. This example has link formatting."""
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename)

        # Simulate custom colour for testing.
        workbook.custom_colors = ['FF0000FF']

        worksheet = workbook.add_worksheet()
        format1 = None
        format2 = workbook.add_format({'color': 'red', 'underline': 1})

        worksheet.write_url('A1', 'http://www.python.org/1', format1)
        worksheet.write_url('A2', 'http://www.python.org/2', format2)

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
