###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

from ...workbook import Workbook
from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("format19.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with unused formats."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        format1 = workbook.add_format({"num_format": "hh:mm;@"})
        format2 = workbook.add_format({"num_format": "hh:mm;@", "bg_color": "yellow"})

        worksheet.write(0, 0, 1, format1)
        worksheet.write(1, 0, 2, format2)

        workbook.close()

        self.assertExcelEqual()
