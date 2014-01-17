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

        filename = 'image20.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.image_dir = test_dir + 'images/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """
        Test the creation of a simple XlsxWriter file with multiple images.
        """
        filename = self.got_filename

        ####################################################

        workbook = Workbook(filename)

        worksheet = workbook.add_worksheet()
        # Add second worksheet for internal link
        workbook.add_worksheet()

        # External link
        url = 'https://www.github.com'
        worksheet.insert_image('C2', self.image_dir + 'train.jpg', url=url)

        url = 'external:./subdir/blank.xlsx'
        worksheet.insert_image('C30', self.image_dir + 'train.jpg', url=url)

        # Internal link
        url = 'internal:Sheet2!A1'
        worksheet.insert_image('C60', self.image_dir + 'train.jpg', url=url)

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
