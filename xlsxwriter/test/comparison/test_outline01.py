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

        filename = 'outline01.xlsx'

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

        worksheet1 = workbook.add_worksheet('Outlined Rows')

        bold = workbook.add_format({'bold': 1})

        worksheet1.set_row(1, None, None, {'level': 2})
        worksheet1.set_row(2, None, None, {'level': 2})
        worksheet1.set_row(3, None, None, {'level': 2})
        worksheet1.set_row(4, None, None, {'level': 2})
        worksheet1.set_row(5, None, None, {'level': 1})

        worksheet1.set_row(6, None, None, {'level': 2})
        worksheet1.set_row(7, None, None, {'level': 2})
        worksheet1.set_row(8, None, None, {'level': 2})
        worksheet1.set_row(9, None, None, {'level': 2})
        worksheet1.set_row(10, None, None, {'level': 1})

        worksheet1.set_column('A:A', 20)

        worksheet1.write('A1', 'Region', bold)
        worksheet1.write('A2', 'North')
        worksheet1.write('A3', 'North')
        worksheet1.write('A4', 'North')
        worksheet1.write('A5', 'North')
        worksheet1.write('A6', 'North Total', bold)

        worksheet1.write('B1', 'Sales', bold)
        worksheet1.write('B2', 1000)
        worksheet1.write('B3', 1200)
        worksheet1.write('B4', 900)
        worksheet1.write('B5', 1200)
        worksheet1.write('B6', '=SUBTOTAL(9,B2:B5)', bold, 4300)

        worksheet1.write('A7', 'South')
        worksheet1.write('A8', 'South')
        worksheet1.write('A9', 'South')
        worksheet1.write('A10', 'South')
        worksheet1.write('A11', 'South Total', bold)

        worksheet1.write('B7', 400)
        worksheet1.write('B8', 600)
        worksheet1.write('B9', 500)
        worksheet1.write('B10', 600)
        worksheet1.write('B11', '=SUBTOTAL(9,B7:B10)', bold, 2100)

        worksheet1.write('A12', 'Grand Total', bold)
        worksheet1.write('B12', '=SUBTOTAL(9,B2:B10)', bold, 6400)

        workbook.close()

        self.assertExcelEqual()
