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

        filename = 'table01.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_2' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        # Ignore increased shared string count.
        self.ignore_files = ['xl/sharedStrings.xml']
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with tables."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_column('C:F', 10.288)

        worksheet.add_table('C3:F13')

        # The following should be ignored since it contains duplicate headers.
        # Ignore the warning.
        import warnings
        warnings.filterwarnings('ignore')

        worksheet.add_table('C3:E3',
                            {'columns': [{'header': 'Column1'},
                                         {'header': 'Column1'}]})
        workbook.close()

        self.assertExcelEqual()
