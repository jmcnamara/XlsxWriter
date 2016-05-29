###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparsion_test import ExcelComparisonTest
from ...workbook import Workbook
from datetime import datetime


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.maxDiff = None

        filename = 'properties04.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file_explicit(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        long_string = 'This is a long string. ' * 11
        long_string += 'AA'

        date = datetime.strptime('2016-12-12 23:00:00', '%Y-%m-%d %H:%M:%S')

        # This test uses explicit property types.
        workbook.set_custom_property('Checked by', 'Adam', 'text')
        workbook.set_custom_property('Date completed', date, 'date')
        workbook.set_custom_property('Document number', 12345, 'number_int')
        workbook.set_custom_property('Reference', 1.2345, 'number')
        workbook.set_custom_property('Source', True, 'bool')
        workbook.set_custom_property('Status', False, 'bool')
        workbook.set_custom_property('Department', long_string, 'text')
        workbook.set_custom_property('Group', 1.23456789012, 'number')

        worksheet.set_column('A:A', 70)
        worksheet.write('A1', "Select 'Office Button -> Prepare -> Properties' to see the file properties.")

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_implicit(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        long_string = 'This is a long string. ' * 11
        long_string += 'AA'

        date = datetime.strptime('2016-12-12 23:00:00', '%Y-%m-%d %H:%M:%S')

        # This test uses implicit property types.
        workbook.set_custom_property('Checked by', 'Adam')
        workbook.set_custom_property('Date completed', date)
        workbook.set_custom_property('Document number', 12345)
        workbook.set_custom_property('Reference', 1.2345)
        workbook.set_custom_property('Source', True)
        workbook.set_custom_property('Status', False)
        workbook.set_custom_property('Department', long_string)
        workbook.set_custom_property('Group', 1.23456789012)

        worksheet.set_column('A:A', 70)
        worksheet.write('A1', "Select 'Office Button -> Prepare -> Properties' to see the file properties.")

        workbook.close()

        self.assertExcelEqual()
