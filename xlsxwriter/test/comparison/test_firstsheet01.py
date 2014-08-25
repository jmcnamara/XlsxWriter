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

        filename = 'firstsheet01.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename)

        worksheet1 = workbook.add_worksheet()
        worksheet2 = workbook.add_worksheet()
        worksheet3 = workbook.add_worksheet()
        worksheet4 = workbook.add_worksheet()
        worksheet5 = workbook.add_worksheet()
        worksheet6 = workbook.add_worksheet()
        worksheet7 = workbook.add_worksheet()
        worksheet8 = workbook.add_worksheet()
        worksheet9 = workbook.add_worksheet()
        worksheet10 = workbook.add_worksheet()
        worksheet11 = workbook.add_worksheet()
        worksheet12 = workbook.add_worksheet()
        worksheet13 = workbook.add_worksheet()
        worksheet14 = workbook.add_worksheet()
        worksheet15 = workbook.add_worksheet()
        worksheet16 = workbook.add_worksheet()
        worksheet17 = workbook.add_worksheet()
        worksheet18 = workbook.add_worksheet()
        worksheet19 = workbook.add_worksheet()
        worksheet20 = workbook.add_worksheet()

        worksheet8.set_first_sheet()
        worksheet20.activate()

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
