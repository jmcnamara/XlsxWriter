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
        self.set_filename("theme02.xlsx")

    def test_create_file(self):
        """Test the addition of a theme file."""

        workbook = Workbook(
            self.got_filename,
            {
                "default_format_properties": {
                    "font_name": "Arial",
                    "font_size": 11,
                    "font_scheme": "minor",
                },
                "default_row_height": 19,
                "default_column_width": 72,
            },
        )

        workbook.use_custom_theme(self.theme_dir + "technic.xml")

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({"type": "line"})

        chart.axis_ids = [85211776, 85262720]

        data = [
            [1, 2, 3, 4, 5],
            [2, 4, 6, 8, 10],
            [3, 6, 9, 12, 15],
        ]

        worksheet.write_column("A1", data[0])
        worksheet.write_column("B1", data[1])
        worksheet.write_column("C1", data[2])

        chart.add_series({"values": "=Sheet1!$A$1:$A$5"})
        chart.add_series({"values": "=Sheet1!$B$1:$B$5"})
        chart.add_series({"values": "=Sheet1!$C$1:$C$5"})

        worksheet.insert_chart("E9", chart)

        workbook.close()

        self.assertExcelEqual()
