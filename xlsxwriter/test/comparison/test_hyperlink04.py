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

        filename = 'hyperlink04.xlsx'

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
        worksheet3 = workbook.add_worksheet('Data Sheet')

        worksheet1.write_url('A1', "internal:Sheet2!A1")
        worksheet1.write_url('A3', "internal:Sheet2!A1:A5")
        worksheet1.write_url('A5', "internal:'Data Sheet'!D5", None, 'Some text')
        worksheet1.write_url('E12', "internal:Sheet1!J1")
        worksheet1.write_url('G17', "internal:Sheet2!A1", None, 'Some text')
        worksheet1.write_url('A18', "internal:Sheet2!A1", None, None, 'Tool Tip 1')
        worksheet1.write_url('A20', "internal:Sheet2!A1", None, 'More text', 'Tool Tip 2')

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_write(self):
        """Test the creation of a simple XlsxWriter file with hyperlinks with write()"""

        workbook = Workbook(self.got_filename)

        # Turn off default URL format for testing.
        workbook.default_url_format = None

        worksheet1 = workbook.add_worksheet()
        worksheet2 = workbook.add_worksheet()
        worksheet3 = workbook.add_worksheet('Data Sheet')

        worksheet1.write('A1', "internal:Sheet2!A1")
        worksheet1.write('A3', "internal:Sheet2!A1:A5")
        worksheet1.write('A5', "internal:'Data Sheet'!D5", None, 'Some text')
        worksheet1.write('E12', "internal:Sheet1!J1")
        worksheet1.write('G17', "internal:Sheet2!A1", None, 'Some text')
        worksheet1.write('A18', "internal:Sheet2!A1", None, None, 'Tool Tip 1')
        worksheet1.write('A20', "internal:Sheet2!A1", None, 'More text', 'Tool Tip 2')

        workbook.close()

        self.assertExcelEqual()
