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
        self.set_filename("format14.xlsx")

    def test_create_file(self):
        """Test the center across format."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        center = workbook.add_format()

        center.set_center_across()

        worksheet.write("A1", "foo", center)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_2(self):
        """Test the center across format."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        center = workbook.add_format({"center_across": True})

        worksheet.write("A1", "foo", center)

        workbook.close()

        self.assertExcelEqual()
