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
        self.set_filename("shared_strings01.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        # Test that control characters and any other single byte characters are
        # handled correctly by the sharedstrings module. We skip chr 34 = " in
        # this test since it isn't encoded by Excel as &quot;.
        chars = list(range(127))
        del chars[34]

        for char in chars:
            worksheet.write_string(char, 0, chr(char))

        workbook.close()

        self.assertExcelEqual()
