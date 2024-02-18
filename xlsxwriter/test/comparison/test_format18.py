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
        self.set_filename("format18.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with a quote prefix."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        quote_prefix = workbook.add_format({"quote_prefix": True})

        worksheet.write_string(0, 0, "= Hello", quote_prefix)

        workbook.close()

        self.assertExcelEqual()
