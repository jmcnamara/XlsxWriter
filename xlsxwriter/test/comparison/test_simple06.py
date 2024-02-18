###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparison_test import ExcelComparisonTest
from datetime import datetime, date, time
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("simple04.xlsx")

    def test_create_file(self):
        """Test dates and times."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_column(0, 0, 12)

        format1 = workbook.add_format({"num_format": 20})
        format2 = workbook.add_format({"num_format": 14})

        date1 = datetime.strptime("12:00", "%H:%M")
        date2 = datetime.strptime("2013-01-27", "%Y-%m-%d")

        worksheet.write_datetime(0, 0, date1, format1)
        worksheet.write_datetime(1, 0, date2, format2)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_write(self):
        """Test dates and times with write() method."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_column(0, 0, 12)

        format1 = workbook.add_format({"num_format": 20})
        format2 = workbook.add_format({"num_format": 14})

        date1 = datetime.strptime("12:00", "%H:%M")
        date2 = datetime.strptime("2013-01-27", "%Y-%m-%d")

        worksheet.write(0, 0, date1, format1)
        worksheet.write(1, 0, date2, format2)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_A1(self):
        """Test dates and times in A1 notation."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_column("A:A", 12)

        format1 = workbook.add_format({"num_format": 20})
        format2 = workbook.add_format({"num_format": 14})

        date1 = datetime.strptime("12:00", "%H:%M")
        date2 = datetime.strptime("2013-01-27", "%Y-%m-%d")

        worksheet.write_datetime("A1", date1, format1)
        worksheet.write_datetime("A2", date2, format2)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_date_and_time1(self):
        """Test dates and times with datetime .date and .time."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_column("A:A", 12)

        format1 = workbook.add_format({"num_format": 20})
        format2 = workbook.add_format({"num_format": 14})

        date1 = time(12)
        date2 = date(2013, 1, 27)

        worksheet.write_datetime("A1", date1, format1)
        worksheet.write_datetime("A2", date2, format2)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_date_and_time2(self):
        """Test dates and times with datetime .date and .time. and write()"""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_column("A:A", 12)

        format1 = workbook.add_format({"num_format": 20})
        format2 = workbook.add_format({"num_format": 14})

        date1 = time(12)
        date2 = date(2013, 1, 27)

        worksheet.write("A1", date1, format1)
        worksheet.write("A2", date2, format2)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_in_memory(self):
        """Test dates and times."""

        workbook = Workbook(self.got_filename, {"in_memory": True})

        worksheet = workbook.add_worksheet()

        worksheet.set_column(0, 0, 12)

        format1 = workbook.add_format({"num_format": 20})
        format2 = workbook.add_format({"num_format": 14})

        date1 = datetime.strptime("12:00", "%H:%M")
        date2 = datetime.strptime("2013-01-27", "%Y-%m-%d")

        worksheet.write_datetime(0, 0, date1, format1)
        worksheet.write_datetime(1, 0, date2, format2)

        workbook.close()

        self.assertExcelEqual()
