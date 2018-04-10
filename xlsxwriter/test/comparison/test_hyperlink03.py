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

        filename = 'hyperlink03.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with hyperlinks."""

        workbook = Workbook(self.got_filename)

        # Turn off default URL format for testing.
        workbook.default_url_format = None

        worksheet1 = workbook.add_worksheet()
        worksheet2 = workbook.add_worksheet()

        worksheet1.write_url('A1', 'http://www.perl.org/')
        worksheet1.write_url('D4', 'http://www.perl.org/')
        worksheet1.write_url('A8', 'http://www.perl.org/')
        worksheet1.write_url('B6', 'http://www.cpan.org/')
        worksheet1.write_url('F12', 'http://www.cpan.org/')

        worksheet2.write_url('C2', 'http://www.google.com/')
        worksheet2.write_url('C5', 'http://www.cpan.org/')
        worksheet2.write_url('C7', 'http://www.perl.org/')

        workbook.close()

        self.assertExcelEqual()
