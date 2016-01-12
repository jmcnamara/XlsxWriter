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

        filename = 'hyperlink12.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with hyperlinks. This example has link formatting."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        format = workbook.add_format({'color': 'blue', 'underline': 1})

        worksheet.write_url('A1', 'mailto:jmcnamara@cpan.org', format)

        worksheet.write_url('A3', 'ftp://perl.org/', format)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_write(self):
        """Test the creation of a simple XlsxWriter file with hyperlinks. This example has link formatting and uses write()"""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        format = workbook.add_format({'color': 'blue', 'underline': 1})

        worksheet.write('A1', 'mailto:jmcnamara@cpan.org', format)

        worksheet.write('A3', 'ftp://perl.org/', format)

        workbook.close()

        self.assertExcelEqual()
