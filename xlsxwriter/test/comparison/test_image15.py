###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

from xlsxwriter.workbook import Workbook

from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("image15.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_row(1, 4.5)
        worksheet.set_row(2, 35.25)
        worksheet.set_column("C:E", 3.29)
        worksheet.set_column("F:F", 10.71)

        worksheet.insert_image(
            "C2", self.image_dir + "logo.png", {"x_offset": 13, "y_offset": 2}
        )

        workbook.close()

        self.assertExcelEqual()
