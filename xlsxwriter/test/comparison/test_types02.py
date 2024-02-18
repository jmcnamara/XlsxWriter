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
        self.set_filename("types02.xlsx")

    def test_write_boolean(self):
        """Test writing boolean."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        worksheet.write_boolean(0, 0, True)
        worksheet.write_boolean(1, 0, False)

        workbook.close()

        self.assertExcelEqual()

    def test_write_boolean_write(self):
        """Test writing boolean with write()."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        worksheet.write(0, 0, True)
        worksheet.write(1, 0, False)

        workbook.close()

        self.assertExcelEqual()
