###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparsion_test import ExcelComparisonTest
from ...workbook import Workbook
from ...compatibility import BytesIO


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.maxDiff = None

        filename = 'header_image03.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.image_dir = test_dir + 'images/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {'xl/worksheets/sheet1.xml': ['<pageMargins', '<pageSetup']}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_header('&L&G&C&G&R&G',
                             {'image_left': self.image_dir + 'red.jpg',
                              'image_center': self.image_dir + 'blue.jpg',
                              'image_right': self.image_dir + 'yellow.jpg'})

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_with_picture(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_header('&L&[Picture]&C&G&R&[Picture]',
                             {'image_left': self.image_dir + 'red.jpg',
                              'image_center': self.image_dir + 'blue.jpg',
                              'image_right': self.image_dir + 'yellow.jpg'})

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_from_bytesio(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        image_file_left = open(self.image_dir + 'red.jpg', 'rb')
        image_data_left = BytesIO(image_file_left.read())
        image_file_left.close()

        image_file_center = open(self.image_dir + 'blue.jpg', 'rb')
        image_data_center = BytesIO(image_file_center.read())
        image_file_center.close()

        image_file_right = open(self.image_dir + 'yellow.jpg', 'rb')
        image_data_right = BytesIO(image_file_right.read())
        image_file_right.close()

        worksheet.set_header('&L&G&C&G&R&G',
                             {'image_left': 'red.jpg',
                              'image_center': 'blue.jpg',
                              'image_right': 'yellow.jpg',
                              'image_data_left': image_data_left,
                              'image_data_center': image_data_center,
                              'image_data_right': image_data_right,
                              })

        workbook.close()

        self.assertExcelEqual()
