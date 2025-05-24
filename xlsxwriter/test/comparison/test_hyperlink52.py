###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

from ...workbook import Workbook
from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):

        self.set_filename("hyperlink52.xlsx")

    def test_create_file(self):
        """
        Test the creation of a simple XlsxWriter file with hyperlinks.
        """

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.write_url(0, 0, "dynamicsnav://www.example.com")

        workbook.close()

        self.assertExcelEqual()
