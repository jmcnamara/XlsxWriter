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
        self.set_filename("excel2003_style01.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename, {"excel2003_style": True})
        workbook.add_worksheet()

        workbook.close()

        self.assertExcelEqual()
