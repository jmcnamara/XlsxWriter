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

        filename = 'table15.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with tables."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        data = [
            ['Foo', 1234, 2000, 4321],
            ['Bar', 1256, 0, 4320],
            ['Baz', 2234, 3000, 4332],
            ['Bop', 1324, 1000, 4333],
        ]

        worksheet.set_column('C:F', 10.288)

        worksheet.add_table('C2:F6', {'data': data})

        workbook.close()

        self.assertExcelEqual()
