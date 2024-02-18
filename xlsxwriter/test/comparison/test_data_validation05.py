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
        self.set_filename("data_validation02.xlsx")

    def test_create_file(self):
        """Test the creation of an  XlsxWriter file data validation."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.data_validation(
            "C2",
            {
                "validate": "list",
                "value": ["Foo", "Bar", "Baz"],
                "input_title": "This is the input title",
                "input_message": "This is the input message",
            },
        )

        # The following should be rejected because the input message is too long.
        input_title = "This is the longest input title1"
        input_message = "This is the longest input message " + ("a" * 222)
        values = [
            "Foobar",
            "Foobas",
            "Foobat",
            "Foobau",
            "Foobav",
            "Foobaw",
            "Foobax",
            "Foobay",
            "Foobaz",
            "Foobba",
            "Foobbb",
            "Foobbc",
            "Foobbd",
            "Foobbe",
            "Foobbf",
            "Foobbg",
            "Foobbh",
            "Foobbi",
            "Foobbj",
            "Foobbk",
            "Foobbl",
            "Foobbm",
            "Foobbn",
            "Foobbo",
            "Foobbp",
            "Foobbq",
            "Foobbr",
            "Foobbs",
            "Foobbt",
            "Foobbu",
            "Foobbv",
            "Foobbw",
            "Foobbx",
            "Foobby",
            "Foobbz",
            "Foobca",
            "End",
        ]

        # Ignore the warnings raised by data_validation().
        import warnings

        warnings.filterwarnings("ignore")

        worksheet.data_validation(
            "D6",
            {
                "validate": "list",
                "value": values,
                "input_title": input_title,
                "input_message": input_message,
            },
        )

        workbook.close()

        self.assertExcelEqual()
