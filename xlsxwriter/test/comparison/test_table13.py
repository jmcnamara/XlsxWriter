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
        self.set_filename("table13.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with tables."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        format1 = workbook.add_format({"num_format": 2, "dxf_index": 2})
        format2 = workbook.add_format({"num_format": 2, "dxf_index": 1})
        format3 = workbook.add_format({"num_format": 2, "dxf_index": 0})

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
                    {"format": format2},
                    {"format": format3},
                ],
            },
        )

        workbook.close()

        self.assertExcelEqual()
