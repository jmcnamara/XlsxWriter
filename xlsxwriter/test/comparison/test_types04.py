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

        filename = 'types04.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_write_url_default(self):
        """Test writing hyperlinks with strings_to_urls on."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()
        red = workbook.add_format({'font_color': 'red'})

        worksheet.write(0, 0, 'http://www.google.com/', red)
        worksheet.write_string(1, 0, 'http://www.google.com/', red)

        workbook.close()

        self.assertExcelEqual()

    def test_write_url_implicit(self):
        """Test writing hyperlinks with strings_to_urls on."""

        workbook = Workbook(self.got_filename, {'strings_to_urls': True})
        worksheet = workbook.add_worksheet()
        red = workbook.add_format({'font_color': 'red'})

        worksheet.write(0, 0, 'http://www.google.com/', red)
        worksheet.write_string(1, 0, 'http://www.google.com/', red)

        workbook.close()

        self.assertExcelEqual()

    def test_write_url_explicit(self):
        """Test writing hyperlinks with strings_to_urls off."""

        workbook = Workbook(self.got_filename, {'strings_to_urls': False})
        worksheet = workbook.add_worksheet()
        red = workbook.add_format({'font_color': 'red'})

        worksheet.write_url(0, 0, 'http://www.google.com/', red)
        worksheet.write(1, 0, 'http://www.google.com/', red)

        workbook.close()

        self.assertExcelEqual()
