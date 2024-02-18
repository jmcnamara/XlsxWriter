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
        self.set_filename("autofit01.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        # Before writing data, nothing to autofit (should not raise)
        worksheet.autofit()

        # Write something that can be autofit
        worksheet.write_string(0, 0, "A")

        # Check for handling default/None width.
        worksheet.set_column("A:A", None)
        worksheet.autofit()

        # Check for handling 0 width.
        worksheet.set_column("A:A", 0)
        worksheet.autofit()

        # Check for handling user defined width. Autofit shouldn't override.
        worksheet.set_column("A:A", 1.57143)
        worksheet.autofit()

        workbook.close()

        self.assertExcelEqual()
