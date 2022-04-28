###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2022, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparison_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):

        self.set_filename('format16.xlsx')

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with a pattern only."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        pattern = workbook.add_format({'pattern': 2})

        worksheet.write('A1', '', pattern)

        workbook.close()

        self.assertExcelEqual()
