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
        self.set_filename("checkbox05.xlsx")

    def test_create_file_with_insert_checkbox(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        worksheet.insert_checkbox("E9", False)

        cell_format = workbook.add_format(
            {
                "font_color": "#9C0006",
                "bg_color": "#FFC7CE",
            }
        )

        worksheet.conditional_format(
            "E9",
            {
                "type": "cell",
                "format": cell_format,
                "criteria": "equal to",
                "value": "FALSE",
            },
        )

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_with_insert_checkbox_and_manual_format(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        cell_format1 = workbook.add_format({"checkbox": True})

        worksheet.insert_checkbox("E9", False, cell_format1)

        cell_format2 = workbook.add_format(
            {
                "font_color": "#9C0006",
                "bg_color": "#FFC7CE",
            }
        )

        worksheet.conditional_format(
            "E9",
            {
                "type": "cell",
                "format": cell_format2,
                "criteria": "equal to",
                "value": "FALSE",
            },
        )

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_with_boolean_and_format(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        cell_format1 = workbook.add_format({"checkbox": True})

        worksheet.write("E9", False, cell_format1)

        cell_format2 = workbook.add_format(
            {
                "font_color": "#9C0006",
                "bg_color": "#FFC7CE",
            }
        )

        worksheet.conditional_format(
            "E9",
            {
                "type": "cell",
                "format": cell_format2,
                "criteria": "equal to",
                "value": "FALSE",
            },
        )

        workbook.close()

        self.assertExcelEqual()

    def test_conditional_format_with_boolean(self):
        """Sub-test for conditional format value as a Python boolean."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        cell_format1 = workbook.add_format({"checkbox": True})

        worksheet.write("E9", False, cell_format1)

        cell_format2 = workbook.add_format(
            {
                "font_color": "#9C0006",
                "bg_color": "#FFC7CE",
            }
        )

        worksheet.conditional_format(
            "E9",
            {
                "type": "cell",
                "format": cell_format2,
                "criteria": "equal to",
                "value": False,
            },
        )

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_with_color_type(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        worksheet.insert_checkbox("E9", False)

        cell_format = workbook.add_format(
            {
                "font_color": Color("#9C0006"),
                "bg_color": Color("#FFC7CE"),
            }
        )

        worksheet.conditional_format(
            "E9",
            {
                "type": "cell",
                "format": cell_format,
                "criteria": "equal to",
                "value": "FALSE",
            },
        )

        workbook.close()

        self.assertExcelEqual()
