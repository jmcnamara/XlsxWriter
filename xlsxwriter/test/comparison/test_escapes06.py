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

        filename = 'escapes06.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file a num format thatrequire XML escaping."""
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename)

        worksheet = workbook.add_worksheet()
        num_format = workbook.add_format({'num_format': '[Red]0.0%\\ "a"'})

        worksheet.set_column('A:A', 14)

        worksheet.write('A1', 123, num_format)

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
