###############################################################################
# _*_ coding: utf-8
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013, John McNamara, jmcnamara@cpan.org
#
from __future__ import unicode_literals
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

        filename = 'utf8_08.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = ['xl/printerSettings/printerSettings1.bin',
                             'xl/worksheets/_rels/sheet1.xml.rels']
        self.ignore_elements = {'[Content_Types].xml': ['<Default Extension="bin"']}

    def test_create_file(self):
        """Test the creation of an XlsxWriter file with utf-8 strings."""
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename)

        worksheet = workbook.add_worksheet()

        worksheet.write('A1', 'Foo')

        worksheet.set_header('&LCafé')
        worksheet.set_footer('&Rclé')

        worksheet.set_paper(9)

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
