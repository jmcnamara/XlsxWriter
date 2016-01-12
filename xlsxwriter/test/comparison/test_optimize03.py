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

        filename = 'optimize03.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename, {'constant_memory': True, 'in_memory': False})
        worksheet = workbook.add_worksheet()

        bold = workbook.add_format({'bold': 1})

        worksheet.set_column('A:A', 36, bold)
        worksheet.set_column('B:B', 20)
        worksheet.set_row(0, 40)

        heading_format = workbook.add_format({
            'bold': 1,
            'font_color': 'blue',
            'font_size': 16,
            'align': 'centre_across',
            'valign': 'vcenter',
        })

        heading_format.text_h_align = 6

        hyperlink_format = workbook.add_format({
            'font_color': 'blue',
            'underline': 1,
        })

        headings = ['Features of Excel::Writer::XLSX', '']
        worksheet.write_row('A1', headings, heading_format)

        text_format = workbook.add_format({
            'bold': 1,
            'italic': 1,
            'font_color': 'red',
            'font_size': 18,
            'font': 'Lucida Calligraphy'
        })

        worksheet.write('A2', "Text")
        worksheet.write('B2', "Hello Excel")
        worksheet.write('A3', "Formatted text")
        worksheet.write('B3', "Hello Excel", text_format)

        num1_format = workbook.add_format({'num_format': '$#,##0.00'})
        num2_format = workbook.add_format({'num_format': ' d mmmm yyy'})

        worksheet.write('A5', "Numbers")
        worksheet.write('B5', 1234.56)
        worksheet.write('A6', "Formatted numbers")
        worksheet.write('B6', 1234.56, num1_format)
        worksheet.write('A7', "Formatted numbers")
        worksheet.write('B7', 37257, num2_format)

        worksheet.write('A8', 'Formulas and functions, "=SIN(PI()/4)"')
        worksheet.write('B8', '=SIN(PI()/4)')

        worksheet.write('A9', "Hyperlinks")
        worksheet.write('B9', 'http://www.perl.com/', hyperlink_format)

        workbook.close()

        self.assertExcelEqual()
