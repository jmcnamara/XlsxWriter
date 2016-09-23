###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparsion_test import ExcelComparisonTest
from datetime import date
from datetime import datetime
from datetime import time
from datetime import timedelta
from datetime import tzinfo
from ...workbook import Workbook


# Simple class to add timezone to dates for testing.
class GMT(tzinfo):

    def utcoffset(self, dt):
        return timedelta(hours=1)

    def dst(self, dt):
        return timedelta(0)

    def tzname(self, dt):
        return "Europe"


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.maxDiff = None

        filename = 'remove_timezone01.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_remove_timezone_none(self):
        """Test write_datetime without timezones."""

        workbook = Workbook(self.got_filename, {'remove_timezone': False})

        worksheet = workbook.add_worksheet()

        worksheet.set_column(0, 0, 20)

        format1 = workbook.add_format({'num_format': 20})
        format2 = workbook.add_format({'num_format': 14})
        format3 = workbook.add_format({'num_format': 22})

        date1 = time(12, 0, 0)
        date2 = date(2016, 9, 23)
        date3 = datetime.strptime('2016-09-12 12:00', "%Y-%m-%d %H:%M")

        worksheet.write_datetime(0, 0, date1, format1)
        worksheet.write_datetime(1, 0, date2, format2)
        worksheet.write_datetime(2, 0, date3, format3)

        workbook.close()

        self.assertExcelEqual()

    def test_remove_timezone_gmt(self):
        """Test write_datetime with timezones."""

        workbook = Workbook(self.got_filename, {'remove_timezone': True})

        worksheet = workbook.add_worksheet()

        worksheet.set_column(0, 0, 20)

        format1 = workbook.add_format({'num_format': 20})
        format2 = workbook.add_format({'num_format': 14})
        format3 = workbook.add_format({'num_format': 22})

        date1 = time(12, 0, 0, tzinfo=GMT())
        date2 = date(2016, 9, 23)
        date3 = datetime.strptime('2016-09-12 12:00', "%Y-%m-%d %H:%M")

        date3 = date3.replace(tzinfo=GMT())

        worksheet.write_datetime(1, 0, date2, format2)
        worksheet.write_datetime(2, 0, date3, format3)

        workbook.close()
