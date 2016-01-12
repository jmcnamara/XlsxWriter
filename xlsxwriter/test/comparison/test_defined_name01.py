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

        filename = 'defined_name01.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = ['xl/printerSettings/printerSettings1.bin',
                             'xl/worksheets/_rels/sheet1.xml.rels']
        self.ignore_elements = {'[Content_Types].xml': ['<Default Extension="bin"'],
                                'xl/worksheets/sheet1.xml': ['<pageMargins', '<pageSetup']}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with defined names."""

        workbook = Workbook(self.got_filename)

        worksheet1 = workbook.add_worksheet()
        worksheet2 = workbook.add_worksheet()
        worksheet3 = workbook.add_worksheet('Sheet 3')

        worksheet1.print_area('A1:E6')
        worksheet1.autofilter('F1:G1')
        worksheet1.write('G1', 'Filter')
        worksheet1.write('F1', 'Auto')
        worksheet1.fit_to_pages(2, 2)

        workbook.define_name("'Sheet 3'!Bar", "='Sheet 3'!$A$1")
        workbook.define_name("Abc", "=Sheet1!$A$1")
        workbook.define_name("Baz", "=0.98")
        workbook.define_name("Sheet1!Bar", "=Sheet1!$A$1")
        workbook.define_name("Sheet2!Bar", "=Sheet2!$A$1")
        workbook.define_name("Sheet2!aaa", "=Sheet2!$A$1")
        workbook.define_name("_Egg", "=Sheet1!$A$1")
        workbook.define_name("_Fog", "=Sheet1!$A$1")

        workbook.close()

        self.assertExcelEqual()
