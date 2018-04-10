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

        filename = 'properties01.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        workbook.set_properties({
            'title': 'This is an example spreadsheet',
            'subject': 'With document properties',
            'author': 'Someone',
            'manager': 'Dr. Heinz Doofenshmirtz',
            'company': 'of Wolves',
            'category': 'Example spreadsheets',
            'keywords': 'Sample, Example, Properties',
            'comments': 'Created with Perl and Excel::Writer::XLSX',
            'status': 'Quo'})

        worksheet.set_column('A:A', 70)
        worksheet.write('A1', "Select 'Office Button -> Prepare -> Properties' to see the file properties.")

        workbook.close()

        self.assertExcelEqual()
