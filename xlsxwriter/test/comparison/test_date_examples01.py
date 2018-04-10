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

        filename = 'date_examples01.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Example spreadsheet used in the tutorial."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        # Widen column A for extra visibility.
        worksheet.set_column('A:A', 30)

        # A number to convert to a date.
        number = 41333.5

        # Write it as a number without formatting.
        worksheet.write('A1', number)  # 41333.5

        format2 = workbook.add_format({'num_format': 'dd/mm/yy'})
        worksheet.write('A2', number, format2)  # 28/02/13

        format3 = workbook.add_format({'num_format': 'mm/dd/yy'})
        worksheet.write('A3', number, format3)  # 02/28/13

        format4 = workbook.add_format({'num_format': 'd\\-m\\-yyyy'})
        worksheet.write('A4', number, format4)  # 28-2-2013

        format5 = workbook.add_format({'num_format': 'dd/mm/yy\\ hh:mm'})
        worksheet.write('A5', number, format5)  # 28/02/13 12:00

        format6 = workbook.add_format({'num_format': 'd\\ mmm\\ yyyy'})
        worksheet.write('A6', number, format6)  # 28 Feb 2013

        format7 = workbook.add_format({'num_format': 'mmm\\ d\\ yyyy\\ hh:mm\\ AM/PM'})
        worksheet.write('A7', number, format7)  # Feb 28 2008 12:00 PM

        workbook.close()

        self.assertExcelEqual()
