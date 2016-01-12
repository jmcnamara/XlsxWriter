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

        filename = 'escapes06.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file a num format thatrequire XML escaping."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        num_format = workbook.add_format({'num_format': '[Red]0.0%\\ "a"'})

        worksheet.set_column('A:A', 14)

        worksheet.write('A1', 123, num_format)

        workbook.close()

        self.assertExcelEqual()
