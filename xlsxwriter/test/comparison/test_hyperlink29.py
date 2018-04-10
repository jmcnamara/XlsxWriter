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

        filename = 'hyperlink29.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with hyperlinks."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        format1 = workbook.add_format({'hyperlink': True})
        format2 = workbook.add_format({'color': 'red', 'underline': 1})

        worksheet.write_url('A1', 'http://www.perl.org/', format1)
        worksheet.write_url('A2', 'http://www.perl.com/', format2)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_with_default_format(self):
        """Test the creation of a simple XlsxWriter file with hyperlinks."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        format2 = workbook.add_format({'color': 'red', 'underline': 1})

        worksheet.write_url('A1', 'http://www.perl.org/')
        worksheet.write_url('A2', 'http://www.perl.com/', format2)

        workbook.close()

        self.assertExcelEqual()
