###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

from xlsxwriter.workbook import Workbook

from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("cond_format28.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with conditionalFormatting."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.write_column("A1", [1, 2, 3, 4, 5])

        # The 3_triangles icon set is only available in Excel 2010 and later.
        worksheet.conditional_format(
            "A1:A5",
            {
                "type": "icon_set",
                "icon_style": "3_triangles",
                "reverse_icons": True,
                "icons_only": True,
            },
        )

        workbook.close()

        self.assertExcelEqual()
