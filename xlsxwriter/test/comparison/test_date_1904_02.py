###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparison_test import ExcelComparisonTest
from datetime import date
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("date_1904_02.xlsx")

    def test_create_file(self):
        """Test the creation of a XlsxWriter file with date times in 1900 and1904 epochs."""

        workbook = Workbook(self.got_filename, {"date_1904": True})

        worksheet = workbook.add_worksheet()
        format1 = workbook.add_format({"num_format": 14})

        worksheet.set_column("A:A", 12)

        worksheet.write_datetime("A1", date(1904, 1, 1), format1)
        worksheet.write_datetime("A2", date(1906, 9, 27), format1)
        worksheet.write_datetime("A3", date(1917, 9, 9), format1)
        worksheet.write_datetime("A4", date(1931, 5, 19), format1)
        worksheet.write_datetime("A5", date(2177, 10, 15), format1)
        worksheet.write_datetime("A6", date(4641, 11, 27), format1)

        workbook.close()

        self.assertExcelEqual()
