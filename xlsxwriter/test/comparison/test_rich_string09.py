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
        self.set_filename("rich_string09.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        bold = workbook.add_format({"bold": 1})
        italic = workbook.add_format({"italic": 1})

        worksheet.write("A1", "Foo", bold)
        worksheet.write("A2", "Bar", italic)
        worksheet.write_rich_string("A3", "a", bold, "bc", "defg")

        # Ignore warnings for the following cases.
        import warnings

        warnings.filterwarnings("ignore")

        # The following has 2 consecutive formats so it should be ignored
        # with a warning.
        worksheet.write_rich_string("A3", "a", bold, bold, "bc", "defg")

        # The following have empty strings and should be ignored with a
        # warning.
        worksheet.write_rich_string("A3", "", bold, "bc", "defg")
        worksheet.write_rich_string("A3", "a", bold, "", "defg")
        worksheet.write_rich_string("A3", "a", bold, "bc", "")

        # The following doesn't have enough fragments/formats and should be
        # ignored with a warning.
        worksheet.write_rich_string("A3", "a")
        worksheet.write_rich_string("A3", "a", bold)
        worksheet.write_rich_string("A3", "a", bold, italic)

        workbook.close()

        self.assertExcelEqual()
