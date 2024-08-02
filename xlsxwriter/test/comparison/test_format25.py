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

        self.set_filename("format25.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with automatic color."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        format1 = workbook.add_format(
            {
                "border_color": "#FF9966",
                "border": 1,
            }
        )

        worksheet.write(2, 2, "", format1)

        workbook.close()

        self.assertExcelEqual()
