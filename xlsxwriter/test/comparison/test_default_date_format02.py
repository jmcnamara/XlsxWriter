###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparsion_test import ExcelComparisonTest
from datetime import datetime
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.maxDiff = None

        filename = 'default_date_format02.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file_user_date_format(self):
        """Test write_datetime with explicit date format."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_column(0, 0, 12)

        format1 = workbook.add_format({'num_format': 'dd\\ mm\\ yy'})

        date1 = datetime.strptime('2013-07-25', "%Y-%m-%d")

        worksheet.write_datetime(0, 0, date1, format1)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_default_date_format(self):
        """Test write_datetime with default date format."""

        workbook = Workbook(self.got_filename, {'default_date_format': 'dd\\ mm\\ yy'})

        worksheet = workbook.add_worksheet()

        worksheet.set_column(0, 0, 12)

        date1 = datetime.strptime('2013-07-25', "%Y-%m-%d")

        worksheet.write_datetime(0, 0, date1)

        workbook.close()

        self.assertExcelEqual()
