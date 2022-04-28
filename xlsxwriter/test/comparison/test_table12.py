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

        self.set_filename('table12.xlsx')

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with tables."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        data = [
            ['Foo', 1234, 2000],
            ['Bar', 1256, 4000],
            ['Baz', 2234, 3000],
        ]

        worksheet.set_column('C:F', 10.288)

        worksheet.add_table('C2:F6', {'data': data})

        workbook.close()

        self.assertExcelEqual()
