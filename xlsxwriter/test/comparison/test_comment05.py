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
        self.set_filename("comment05.xlsx")

        # It takes about 50 times longer to run this test with this file
        # included. Only turn it on for pre-release testing.
        self.ignore_files = ["xl/drawings/vmlDrawing1.vml"]

    def test_create_file(self):
        """
        Test the creation of a simple XlsxWriter file with comments.
        Test the VML data and shape ids for blocks of comments > 1024.
        """

        workbook = Workbook(self.got_filename)

        worksheet1 = workbook.add_worksheet()
        worksheet2 = workbook.add_worksheet()
        worksheet3 = workbook.add_worksheet()

        for row in range(0, 127 + 1):
            for col in range(0, 15 + 1):
                worksheet1.write_comment(row, col, "Some text")

        worksheet3.write_comment("A1", "More text")

        worksheet1.set_comments_author("John")
        worksheet3.set_comments_author("John")

        workbook.close()

        self.assertExcelEqual()
