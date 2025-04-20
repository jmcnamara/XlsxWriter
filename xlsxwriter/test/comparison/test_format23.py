###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

from xlsxwriter.color import Color
from xlsxwriter.workbook import Workbook

from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("format23.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with automatic color."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        format1 = workbook.add_format(
            {
                "font_color": "automatic",
                "diag_border": 1,
                "diag_type": 2,
                "diag_color": "automatic",
            }
        )

        worksheet.write(0, 0, "Foo", format1)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_with_color_type1(self):
        """Test the creation of a simple XlsxWriter file with automatic color."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        format1 = workbook.add_format(
            {
                "font_color": "automatic",
                "diag_border": 1,
                "diag_type": 2,
                "diag_color": Color("automatic"),
            }
        )

        worksheet.write(0, 0, "Foo", format1)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_with_color_type2(self):
        """Test the creation of a simple XlsxWriter file with automatic color."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        format1 = workbook.add_format(
            {
                "font_color": "automatic",
                "diag_border": 1,
                "diag_type": 2,
                "diag_color": Color.automatic(),
            }
        )

        worksheet.write(0, 0, "Foo", format1)

        workbook.close()

        self.assertExcelEqual()
