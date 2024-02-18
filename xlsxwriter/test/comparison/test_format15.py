##############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparison_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("format15.xlsx")

    def test_create_file_zero_number_format(self):
        """Test the creation of a simple XlsxWriter file 0 number format."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        format1 = workbook.add_format({"bold": 1})
        format2 = workbook.add_format({"bold": 1, "num_format": 0})

        worksheet.write("A1", 1, format1)
        worksheet.write("A2", 2, format2)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_zero_number_format_string(self):
        """Test the creation of a simple XlsxWriter file 0 number format."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        format1 = workbook.add_format({"bold": 1})
        format2 = workbook.add_format({"bold": 1, "num_format": "0"})

        worksheet.write("A1", 1, format1)
        worksheet.write("A2", 2, format2)

        workbook.close()

        self.assertExcelEqual()
