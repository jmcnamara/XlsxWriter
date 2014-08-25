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

        filename = 'types05.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = ['xl/calcChain.xml',
                             '[Content_Types].xml',
                             'xl/_rels/workbook.xml.rels']
        self.ignore_elements = {}

    def test_write_formula_default(self):
        """Test writing formulas with strings_to_formulas on."""
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename)
        worksheet = workbook.add_worksheet()

        worksheet.write(0, 0, '=1+1', None, 2)
        worksheet.write_string(1, 0, '=1+1')

        workbook.close()

        ####################################################

        got, exp = _compare_xlsx_files(self.got_filename,
                                       self.exp_filename,
                                       self.ignore_files,
                                       self.ignore_elements)

        self.assertEqual(got, exp)

    def test_write_formula_implicit(self):
        """Test writing formulas with strings_to_formulas on."""
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename, {'strings_to_formulas': True})
        worksheet = workbook.add_worksheet()

        worksheet.write(0, 0, '=1+1', None, 2)
        worksheet.write_string(1, 0, '=1+1')

        workbook.close()

        ####################################################

        got, exp = _compare_xlsx_files(self.got_filename,
                                       self.exp_filename,
                                       self.ignore_files,
                                       self.ignore_elements)

        self.assertEqual(got, exp)

    def test_write_formula_explicit(self):
        """Test writing formulas with strings_to_formulas off."""
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename, {'strings_to_formulas': False})
        worksheet = workbook.add_worksheet()

        worksheet.write_formula(0, 0, '=1+1', None, 2)
        worksheet.write(1, 0, '=1+1')

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
