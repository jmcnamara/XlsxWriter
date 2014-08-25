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

        filename = 'simple05.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test font formatting."""
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_row(5, 18)
        worksheet.set_row(6, 18)

        format1 = workbook.add_format({'bold': 1})
        format2 = workbook.add_format({'italic': 1})
        format3 = workbook.add_format({'bold': 1, 'italic': 1})
        format4 = workbook.add_format({'underline': 1})
        format5 = workbook.add_format({'font_strikeout': 1})
        format6 = workbook.add_format({'font_script': 1})
        format7 = workbook.add_format({'font_script': 2})

        worksheet.write_string(0, 0, 'Foo', format1)
        worksheet.write_string(1, 0, 'Foo', format2)
        worksheet.write_string(2, 0, 'Foo', format3)
        worksheet.write_string(3, 0, 'Foo', format4)
        worksheet.write_string(4, 0, 'Foo', format5)
        worksheet.write_string(5, 0, 'Foo', format6)
        worksheet.write_string(6, 0, 'Foo', format7)

        workbook.close()

        ####################################################

        got, exp = _compare_xlsx_files(self.got_filename,
                                       self.exp_filename,
                                       self.ignore_files,
                                       self.ignore_elements)

        self.assertEqual(got, exp)

    def test_create_file_in_memory(self):
        """Test font formatting."""
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename, {'in_memory': True})

        worksheet = workbook.add_worksheet()

        worksheet.set_row(5, 18)
        worksheet.set_row(6, 18)

        format1 = workbook.add_format({'bold': 1})
        format2 = workbook.add_format({'italic': 1})
        format3 = workbook.add_format({'bold': 1, 'italic': 1})
        format4 = workbook.add_format({'underline': 1})
        format5 = workbook.add_format({'font_strikeout': 1})
        format6 = workbook.add_format({'font_script': 1})
        format7 = workbook.add_format({'font_script': 2})

        worksheet.write_string(0, 0, 'Foo', format1)
        worksheet.write_string(1, 0, 'Foo', format2)
        worksheet.write_string(2, 0, 'Foo', format3)
        worksheet.write_string(3, 0, 'Foo', format4)
        worksheet.write_string(4, 0, 'Foo', format5)
        worksheet.write_string(5, 0, 'Foo', format6)
        worksheet.write_string(6, 0, 'Foo', format7)

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
