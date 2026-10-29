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
        self.set_filename("cond_format26.xlsx")

    def test_create_file(self):
        """Test a conditional format pattern fill without a background color."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        cell_format = workbook.add_format({"pattern": 2, "fg_color": "#FF0000"})

        worksheet.write("A1", 10)
        worksheet.write("A2", 20)
        worksheet.write("A3", 30)
        worksheet.write("A4", 40)

        worksheet.conditional_format(
            "A1:A4",
            {
                "type": "cell",
                "criteria": ">",
                "value": 25,
                "format": cell_format,
            },
        )

        workbook.close()

        self.assertExcelEqual()
