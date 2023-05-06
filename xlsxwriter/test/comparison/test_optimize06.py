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
        self.set_filename("optimize06.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(
            self.got_filename, {"constant_memory": True, "in_memory": False}
        )
        worksheet = workbook.add_worksheet()

        # Test that control characters and any other single byte characters are
        # handled correctly by the SharedStrings module. We skip chr 34 = " in
        # this test since it isn't encoded by Excel as &quot;.
        ordinals = list(range(0, 34))
        ordinals.extend(range(35, 128))

        for i in ordinals:
            worksheet.write_string(i, 0, chr(i))

        workbook.close()

        self.assertExcelEqual()
