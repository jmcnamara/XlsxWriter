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

        filename = 'simple09.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        # Test data out of range. These should be ignored.
        worksheet.write('A0', 'foo')
        worksheet.write(-1, -1, 'foo')
        worksheet.write(0, -1, 'foo')
        worksheet.write(-1, 0, 'foo')
        worksheet.write(1048576, 0, 'foo')
        worksheet.write(0, 16384, 'foo')

        workbook.close()

        self.assertExcelEqual()
