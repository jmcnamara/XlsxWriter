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

        filename = 'format01.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_members = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with unused formats."""
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename)

        worksheet1 = workbook.add_worksheet()
        worksheet2 = workbook.add_worksheet('Data Sheet')
        worksheet3 = workbook.add_worksheet()

        unused1 = workbook.add_format({'bold': 1})
        bold = workbook.add_format({'bold': 1})
        unused2 = workbook.add_format({'bold': 1})
        unused3 = workbook.add_format({'italic': 1})

        worksheet1.write('A1', 'Foo')
        worksheet1.write('A2', 123)

        worksheet3.write('B2', 'Foo')
        worksheet3.write('B3', 'Bar', bold)
        worksheet3.write('C4', 234)

        workbook.close()

        ####################################################

        exp, got = _compare_xlsx_files(self.got_filename,
                                       self.exp_filename,
                                       self.ignore_members,
                                       self.ignore_elements)

        self.assertEqual(got, exp)

    def tearDown(self):
        # Cleanup.
        if os.path.exists(self.got_filename):
            os.remove(self.got_filename)


if __name__ == '__main__':
    unittest.main()
