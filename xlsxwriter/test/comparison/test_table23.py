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

        filename = 'table23.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with tables that include metacharacters in their column names."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        data = [[1, 2], [3, 4]]

        worksheet.add_table('A1:B3',
            {
                'data': data,
                'total_row': True,
                'columns': [
                    { 'header': 'meta # chars', 'total_function': 'sum'},
                    { 'header': 'this col is fine', 'total_function': 'sum'},
                ],
            }
        )

        workbook.close()

        self.assertExcelEqual()
