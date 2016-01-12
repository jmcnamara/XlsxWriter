###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparsion_test import ExcelComparisonTest
from ...workbook import Workbook
from ...sharedstrings import SharedStringTable


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.maxDiff = None

        filename = 'hyperlink19.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = ['xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels']
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with hyperlinks."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.write_url('A1', 'http://www.perl.com/')

        # Maintain the link but overwrite string with a formula.
        worksheet.write_formula('A1', '=1+1', None, 2)

        # Reset the SST for testing.
        workbook.str_table = SharedStringTable()

        workbook.close()

        self.assertExcelEqual()
