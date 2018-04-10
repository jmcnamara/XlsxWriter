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

        filename = 'outline05.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = ['xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels']
        self.ignore_elements = {}

    def test_create_file(self):
        """
        Test the creation of a outlines in a XlsxWriter file. These tests are
        based on the outline programs in the examples directory.
        """

        workbook = Workbook(self.got_filename)

        worksheet2 = workbook.add_worksheet('Collapsed Rows')

        bold = workbook.add_format({'bold': 1})

        worksheet2.set_row(1, None, None, {'level': 2, 'hidden': True})
        worksheet2.set_row(2, None, None, {'level': 2, 'hidden': True})
        worksheet2.set_row(3, None, None, {'level': 2, 'hidden': True})
        worksheet2.set_row(4, None, None, {'level': 2, 'hidden': True})
        worksheet2.set_row(5, None, None, {'level': 1, 'hidden': True, 'collapsed': True})

        worksheet2.set_row(6, None, None, {'level': 2, 'hidden': True})
        worksheet2.set_row(7, None, None, {'level': 2, 'hidden': True})
        worksheet2.set_row(8, None, None, {'level': 2, 'hidden': True})
        worksheet2.set_row(9, None, None, {'level': 2, 'hidden': True})
        worksheet2.set_row(10, None, None, {'level': 1, 'hidden': True, 'collapsed': True})

        worksheet2.set_row(11, None, None, {'collapsed': True})

        worksheet2.set_column('A:A', 20)
        worksheet2.set_selection('A14')

        worksheet2.write('A1', 'Region', bold)
        worksheet2.write('A2', 'North')
        worksheet2.write('A3', 'North')
        worksheet2.write('A4', 'North')
        worksheet2.write('A5', 'North')
        worksheet2.write('A6', 'North Total', bold)

        worksheet2.write('B1', 'Sales', bold)
        worksheet2.write('B2', 1000)
        worksheet2.write('B3', 1200)
        worksheet2.write('B4', 900)
        worksheet2.write('B5', 1200)
        worksheet2.write('B6', '=SUBTOTAL(9,B2:B5)', bold, 4300)

        worksheet2.write('A7', 'South')
        worksheet2.write('A8', 'South')
        worksheet2.write('A9', 'South')
        worksheet2.write('A10', 'South')
        worksheet2.write('A11', 'South Total', bold)

        worksheet2.write('B7', 400)
        worksheet2.write('B8', 600)
        worksheet2.write('B9', 500)
        worksheet2.write('B10', 600)
        worksheet2.write('B11', '=SUBTOTAL(9,B7:B10)', bold, 2100)

        worksheet2.write('A12', 'Grand Total', bold)
        worksheet2.write('B12', '=SUBTOTAL(9,B2:B10)', bold, 6400)

        workbook.close()

        self.assertExcelEqual()
