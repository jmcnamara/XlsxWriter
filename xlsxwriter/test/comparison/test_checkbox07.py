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
        self.set_filename("checkbox07.xlsx")

    def test_create_file_with_checkboxes_in_table(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()
        cell_format = workbook.add_format({"checkbox": True})

        data = [
            [1, True],
            [2, False],
            [3, False],
            [4, True],
        ]

        worksheet.add_table(
            "A1:B5",
            {
                "data": data,
                "columns": [
                    {"header": "Col1"},
                    {"header": "Col2", "format": cell_format},
                ],
            },
        )

        workbook.close()

        self.assertExcelEqual()
