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

        filename = 'image01.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.image_dir = test_dir + 'images/'
        self.got_filename = test_dir + '_test_bytes_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        image_file = open(self.image_dir + 'red.png', 'rb')
        image_data = BytesIO(image_file.read())
        image_file.close()

        worksheet.insert_image('E9', 'red.png', {'image_data': image_data})

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_in_memory(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename, {'in_memory': True})

        worksheet = workbook.add_worksheet()

        image_file = open(self.image_dir + 'red.png', 'rb')
        image_data = BytesIO(image_file.read())
        image_file.close()

        worksheet.insert_image('E9', 'red.png', {'image_data': image_data})

        workbook.close()

        self.assertExcelEqual()
