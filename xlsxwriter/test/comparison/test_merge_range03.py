###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2023, John McNamara, jmcnamara@cpan.org
#

from ...workbook import Workbook
from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("merge_range03.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        format = workbook.add_format({"align": "center"})

        worksheet.merge_range(1, 1, 1, 2, "Foo", format)
        worksheet.merge_range(1, 3, 1, 4, "Foo", format)
        worksheet.merge_range(1, 5, 1, 6, "Foo", format)

        workbook.close()

        self.assertExcelEqual()
