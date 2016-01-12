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

        filename = 'hyperlink20.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_hyperlink_formating_explicit(self):
        """Test the creation of a simple XlsxWriter file with hyperlinks. This example has link formatting."""

        workbook = Workbook(self.got_filename)

        # Simulate custom colour for testing.
        workbook.custom_colors = ['FF0000FF']

        worksheet = workbook.add_worksheet()
        format1 = workbook.add_format({'color': 'blue', 'underline': 1})
        format2 = workbook.add_format({'color': 'red', 'underline': 1})

        worksheet.write_url('A1', 'http://www.python.org/1', format1)
        worksheet.write_url('A2', 'http://www.python.org/2', format2)

        workbook.close()

        self.assertExcelEqual()

    def test_hyperlink_formating_implicit(self):
        """Test the creation of a simple XlsxWriter file with hyperlinks. This example has link formatting."""

        workbook = Workbook(self.got_filename)

        # Simulate custom colour for testing.
        workbook.custom_colors = ['FF0000FF']

        worksheet = workbook.add_worksheet()
        format1 = None
        format2 = workbook.add_format({'color': 'red', 'underline': 1})

        worksheet.write_url('A1', 'http://www.python.org/1', format1)
        worksheet.write_url('A2', 'http://www.python.org/2', format2)

        workbook.close()

        self.assertExcelEqual()
