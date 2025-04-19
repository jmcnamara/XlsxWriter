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
        self.set_filename("hyperlink11.xlsx")

    def test_link_format_explicit(self):
        """
        Test the creation of a simple XlsxWriter file with hyperlinks. This
        example has link formatting.

        """

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        url_format = workbook.add_format({"font_color": "blue", "underline": 1})

        worksheet.write_url("A1", "http://www.perl.org/", url_format)

        workbook.close()

        self.assertExcelEqual()
