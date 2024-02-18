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
        self.set_filename("format11.xlsx")

    def test_create_file(self):
        """Test a vertical and horizontal centered format."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        centered = workbook.add_format({"align": "center", "valign": "vcenter"})

        worksheet.write("B2", "Foo", centered)

        workbook.close()

        self.assertExcelEqual()
