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

        filename = 'escapes01.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = ['xl/calcChain.xml',
                             '[Content_Types].xml',
                             'xl/_rels/workbook.xml.rels']
        self.ignore_elements = {}

    def test_create_file(self):
        """Test creation of a file with strings that require XML escaping."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet('5&4')

        worksheet.write_formula(0, 0, '=IF(1>2,0,1)', None, 1)
        worksheet.write_formula(1, 0, """=CONCATENATE("'","<>&")""", None, "'<>&")
        worksheet.write_formula(2, 0, '=1&"b"', None, '1b')
        worksheet.write_formula(3, 0, """="'\"""", None, "'")
        worksheet.write_formula(4, 0, '=""""', None, '"')
        worksheet.write_formula(5, 0, '="&" & "&"', None, '&&')

        worksheet.write_string(7, 0, '"&<>')

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_write(self):
        """Test formulas with write() method."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet('5&4')

        worksheet.write(0, 0, '=IF(1>2,0,1)', None, 1)
        worksheet.write(1, 0, """=CONCATENATE("'","<>&")""", None, "'<>&")
        worksheet.write(2, 0, '=1&"b"', None, '1b')
        worksheet.write(3, 0, """="'\"""", None, "'")
        worksheet.write(4, 0, '=""""', None, '"')
        worksheet.write(5, 0, '="&" & "&"', None, '&&')

        worksheet.write_string(7, 0, '"&<>')

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_A1(self):
        """Test formulas with A1 notation."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet('5&4')

        worksheet.write_formula('A1', '=IF(1>2,0,1)', None, 1)
        worksheet.write_formula('A2', """=CONCATENATE("'","<>&")""", None, "'<>&")
        worksheet.write_formula('A3', '=1&"b"', None, '1b')
        worksheet.write_formula('A4', """="'\"""", None, "'")
        worksheet.write_formula('A5', '=""""', None, '"')
        worksheet.write_formula('A6', '="&" & "&"', None, '&&')

        worksheet.write_string(7, 0, '"&<>')

        workbook.close()

        self.assertExcelEqual()
