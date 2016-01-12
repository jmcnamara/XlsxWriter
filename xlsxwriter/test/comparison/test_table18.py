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

        filename = 'table18.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with tables."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        text_wrap = workbook.add_format({'text_wrap': 1})

        worksheet.set_column('C:F', 10.288)
        worksheet.set_row(2, 39)

        worksheet.add_table('C3:F13',
                            {'columns': [{},
                                         {},
                                         {},
                                         {'header': "Column\n4",
                                          'header_format': text_wrap}]})

        worksheet.write('A16', 'hello')

        workbook.close()

        self.assertExcelEqual()
