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
        self.set_filename("escapes06.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file a num format that require XML escaping."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        num_format = workbook.add_format({"num_format": '[Red]0.0%\\ "a"'})

        worksheet.set_column("A:A", 14)

        worksheet.write("A1", 123, num_format)

        workbook.close()

        self.assertExcelEqual()
