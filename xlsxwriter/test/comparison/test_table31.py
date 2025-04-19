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
        self.set_filename("table31.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with tables."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        format1 = workbook.add_format(
            {"bg_color": "#FFFF00", "fg_color": "#FF0000", "pattern": 6}
        )

        data = [
            ["Foo", 1234, 2000, 4321],
            ["Bar", 1256, 4000, 4320],
            ["Baz", 2234, 3000, 4332],
            ["Bop", 1324, 1000, 4333],
        ]

        worksheet.set_column("C:F", 10.288)

        worksheet.add_table(
            "C2:F6",
            {
                "data": data,
                "columns": [
                    {},
                    {"format": format1},
                ],
            },
        )

        workbook.close()

        self.assertExcelEqual()
