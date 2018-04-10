###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparsion_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.maxDiff = None

        filename = 'table08.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = ['xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels']
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with tables."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_column('C:F', 10.288)

        worksheet.write_string('A1', 'Column1')
        worksheet.write_string('B1', 'Column2')
        worksheet.write_string('C1', 'Column3')
        worksheet.write_string('D1', 'Column4')
        worksheet.write_string('E1', 'Total')

        worksheet.add_table('C3:F14', {'total_row': 1,
                                       'columns': [{'total_string': 'Total'},
                                                   {},
                                                   {},
                                                   {'total_function': 'count'},
                                                   ]})

        workbook.close()

        self.assertExcelEqual()
