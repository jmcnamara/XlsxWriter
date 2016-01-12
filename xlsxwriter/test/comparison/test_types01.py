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

        filename = 'types01.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_write_number_as_text(self):
        """Test writing numbers as text."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        worksheet.write_string(0, 0, 'Hello')
        worksheet.write_string(1, 0, '123')

        workbook.close()

        self.assertExcelEqual()

    def test_write_number_as_text_with_write(self):
        """Test writing numbers as text using write() without conversion."""

        workbook = Workbook(self.got_filename, {'strings_to_numbers': False})
        worksheet = workbook.add_worksheet()

        worksheet.write(0, 0, 'Hello')
        worksheet.write(1, 0, '123')

        workbook.close()

        self.assertExcelEqual()
