###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

from datetime import date
from ..excel_comparison_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("autofit05.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        date_format = workbook.add_format({"num_format": 14})

        worksheet.write_datetime(0, 0, date(2023, 1, 1), date_format)
        worksheet.write_datetime(0, 1, date(2023, 12, 12), date_format)

        worksheet.autofit()

        workbook.close()

        self.assertExcelEqual()
