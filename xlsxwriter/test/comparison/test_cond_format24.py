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
        self.set_filename("cond_format24.xlsx")

    def test_create_file(self):
        """Test text conditional formats with double quotes and non-ASCII strings."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        cell_format = workbook.add_format({"bg_color": "#FFFF00"})

        worksheet.write("A1", 'say "hi"')
        worksheet.write("A2", "café")
        worksheet.write("A3", "a😀")
        worksheet.write("A4", "other")

        worksheet.conditional_format(
            "A1:A4",
            {
                "type": "text",
                "criteria": "containing",
                "value": 'say "hi"',
                "format": cell_format,
            },
        )

        worksheet.conditional_format(
            "A1:A4",
            {
                "type": "text",
                "criteria": "begins with",
                "value": "café",
                "format": cell_format,
            },
        )

        worksheet.conditional_format(
            "A1:A4",
            {
                "type": "text",
                "criteria": "ends with",
                "value": "a😀",
                "format": cell_format,
            },
        )

        workbook.close()

        self.assertExcelEqual()
