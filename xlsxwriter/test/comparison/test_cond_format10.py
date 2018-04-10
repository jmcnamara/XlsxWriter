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

        filename = 'cond_format10.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with conditional formatting."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        format1 = workbook.add_format({'bold': 1, 'italic': 1})

        worksheet.write('A1', 'Hello', format1)

        worksheet.write('B3', 10)
        worksheet.write('B4', 20)
        worksheet.write('B5', 30)
        worksheet.write('B6', 40)

        worksheet.conditional_format('B3:B6',
                                     {'type': 'cell',
                                      'format': format1,
                                      'criteria': 'greater than',
                                      'value': 20
                                      })

        workbook.close()

        self.assertExcelEqual()
