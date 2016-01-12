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

        filename = 'set_start_page03.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {'xl/worksheets/sheet1.xml': ['<pageMargins']}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with printer settings."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_start_page(101)
        worksheet.set_paper(9)

        worksheet.vertical_dpi = 200

        worksheet.write('A1', 'Foo')

        workbook.close()

        self.assertExcelEqual()
