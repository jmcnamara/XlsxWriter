###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparsion_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.maxDiff = None

        filename = 'button07.xlsm'

        test_dir = 'xlsxwriter/test/comparison/'
        self.vba_dir = test_dir + 'xlsx_files/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        workbook.set_vba_name()
        worksheet.set_vba_name()

        worksheet.insert_button('C2', {'macro': 'say_hello',
                                       'caption': 'Hello'})

        workbook.add_vba_project(self.vba_dir + 'vbaProject02.bin')

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_explicit_vba_names(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        workbook.set_vba_name('ThisWorkbook')
        worksheet.set_vba_name('Sheet1')

        worksheet.insert_button('C2', {'macro': 'say_hello',
                                       'caption': 'Hello'})

        workbook.add_vba_project(self.vba_dir + 'vbaProject02.bin')

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_implicit_vba_names(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.insert_button('C2', {'macro': 'say_hello',
                                       'caption': 'Hello'})

        workbook.add_vba_project(self.vba_dir + 'vbaProject02.bin')

        workbook.close()

        self.assertExcelEqual()
