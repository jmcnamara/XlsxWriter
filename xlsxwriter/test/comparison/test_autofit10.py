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
        self.set_filename("autofit10.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        bold = workbook.add_format({"bold": True})

        worksheet.write_rich_string(
            0, 0, "F", bold, "o", "o", bold, "b", "a", bold, "r"
        )

        worksheet.write(1, 0, "Bar", bold)

        worksheet.autofit()

        workbook.close()

        self.assertExcelEqual()