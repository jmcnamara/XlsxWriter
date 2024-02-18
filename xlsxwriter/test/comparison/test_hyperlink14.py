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
        self.set_filename("hyperlink14.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with hyperlinks. This example has writes a url in a range."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        format = workbook.add_format({"align": "center"})

        worksheet.merge_range(
            "C4:E5",
            "",
            format,
        )
        worksheet.write_url("C4", "http://www.perl.org/", format, "Perl Home")

        workbook.close()

        self.assertExcelEqual()
