###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2019, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparsion_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):

        self.set_filename('image20.xlsx')

    def test_create_file(self):
        """
        Test the creation of a simple XlsxWriter file with multiple images.
        """

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        # Add second worksheet for internal link
        workbook.add_worksheet()

        # External link
        options = {'url': 'https://www.github.com'}
        worksheet.insert_image('C2', self.image_dir + 'train.jpg', options)

        options = {'url': 'external:./subdir/blank.xlsx'}
        worksheet.insert_image('C30', self.image_dir + 'train.jpg', options)

        # Internal link
        options = {'url': 'internal:Sheet2!A1'}
        worksheet.insert_image('C60', self.image_dir + 'train.jpg', options)

        workbook.close()

        self.assertExcelEqual()
