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

        filename = 'hyperlink03.xlsx'

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

        # Turn off default URL format for testing.
        workbook.default_url_format = None

        worksheet1 = workbook.add_worksheet()
        worksheet2 = workbook.add_worksheet()

        worksheet1.write_url('A1', 'http://www.perl.org/')
        worksheet1.write_url('D4', 'http://www.perl.org/')
        worksheet1.write_url('A8', 'http://www.perl.org/')
        worksheet1.write_url('B6', 'http://www.cpan.org/')
        worksheet1.write_url('F12', 'http://www.cpan.org/')

        worksheet2.write_url('C2', 'http://www.google.com/')
        worksheet2.write_url('C5', 'http://www.cpan.org/')
        worksheet2.write_url('C7', 'http://www.perl.org/')

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
