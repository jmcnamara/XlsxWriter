###############################################################################
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
        self.set_filename("autofit13.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.write_string(0, 0, "Foo")
        worksheet.write_string(0, 1, "Foo bar")
        worksheet.write_string(0, 2, "Foo bar bar")

        worksheet.autofilter(0, 0, 0, 2)

        # Test autofit with filter headers.
        worksheet.autofit()

        workbook.close()

        self.assertExcelEqual()
