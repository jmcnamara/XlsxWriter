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

        self.set_filename('comment04.xlsx')

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with comments."""

        workbook = Workbook(self.got_filename)

        worksheet1 = workbook.add_worksheet()
        worksheet2 = workbook.add_worksheet()
        worksheet3 = workbook.add_worksheet()

        worksheet1.write('A1', 'Foo')
        worksheet1.write_comment('B2', 'Some text')

        worksheet3.write('A1', 'Bar')
        worksheet3.write_comment('C7', 'More text')

        worksheet1.set_comments_author('John')
        worksheet3.set_comments_author('John')

        workbook.close()

        self.assertExcelEqual()
