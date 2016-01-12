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

        filename = 'types07.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = ['xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels']
        self.ignore_elements = {}

    def test_write_nan_and_inf(self):
        """Test writing special numbers."""

        workbook = Workbook(self.got_filename, {'nan_inf_to_errors': True})
        worksheet = workbook.add_worksheet()

        worksheet.write(0, 0, float('nan'))
        worksheet.write(1, 0, float('inf'))
        worksheet.write(2, 0, float('-inf'))

        workbook.close()

        self.assertExcelEqual()

    def test_write_nan_and_inf_write_number(self):
        """Test writing special numbers."""

        workbook = Workbook(self.got_filename, {'nan_inf_to_errors': True})
        worksheet = workbook.add_worksheet()

        worksheet.write_number(0, 0, float('nan'))
        worksheet.write_number(1, 0, float('inf'))
        worksheet.write_number(2, 0, float('-inf'))

        workbook.close()

        self.assertExcelEqual()

    def test_write_nan_and_inf_write_as_string(self):
        """Test writing special numbers."""

        workbook = Workbook(self.got_filename, {'nan_inf_to_errors': True,
                                                'strings_to_numbers': True})
        worksheet = workbook.add_worksheet()

        worksheet.write(0, 0, 'nan')
        worksheet.write(1, 0, 'inf')
        worksheet.write(2, 0, '-inf')

        workbook.close()

        self.assertExcelEqual()
