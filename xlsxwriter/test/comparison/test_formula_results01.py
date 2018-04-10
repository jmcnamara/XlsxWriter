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

        filename = 'formula_results01.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = ['xl/calcChain.xml',
                             '[Content_Types].xml',
                             'xl/_rels/workbook.xml.rels']
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with formula errors."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.write_formula('A1', '1+1', None, 2)
        worksheet.write_formula('A2', '"Foo"', None, 'Foo')
        worksheet.write_formula('A3', 'IF(B3,FALSE,TRUE)', None, True)
        worksheet.write_formula('A4', 'IF(B4,TRUE,FALSE)', None, False)
        worksheet.write_formula('A5', '#DIV/0!', None, '#DIV/0!')
        worksheet.write_formula('A6', '#N/A', None, '#N/A')
        worksheet.write_formula('A7', '#NAME?', None, '#NAME?')
        worksheet.write_formula('A8', '#NULL!', None, '#NULL!')
        worksheet.write_formula('A9', '#NUM!', None, '#NUM!')
        worksheet.write_formula('A10', '#REF!', None, '#REF!')
        worksheet.write_formula('A11', '#VALUE!', None, '#VALUE!')
        worksheet.write_formula('A12', '1/0', None, '#DIV/0!')

        workbook.close()

        self.assertExcelEqual()
