###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

from xlsxwriter.url import Url
from xlsxwriter.workbook import Workbook

from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("hyperlink05.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with hyperlinks."""

        workbook = Workbook(self.got_filename)

        # Turn off default URL format for testing.
        workbook.default_url_format = None

        worksheet = workbook.add_worksheet()

        worksheet.write_url("A1", "http://www.perl.org/")
        worksheet.write_url("A3", "http://www.perl.org/", None, "Perl home")
        worksheet.write_url("A5", "http://www.perl.org/", None, "Perl home", "Tool Tip")
        worksheet.write_url("A7", "http://www.cpan.org/", None, "CPAN", "Download")

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_with_url_type(self):
        """Test the creation of a simple XlsxWriter using Url class"""

        workbook = Workbook(self.got_filename)

        # Turn off default URL format for testing.
        workbook.default_url_format = None

        worksheet = workbook.add_worksheet()

        worksheet.write_url("A1", Url("http://www.perl.org/"))

        url = Url("http://www.perl.org/")
        url.text = "Perl home"
        worksheet.write_url("A3", url)

        url = Url("http://www.perl.org/")
        url.text = "Perl home"
        url.tip = "Tool Tip"
        worksheet.write_url("A5", url)

        url = Url("http://www.cpan.org/")
        url.text = "CPAN"
        url.tip = "Download"
        worksheet.write_url("A7", url)

        workbook.close()

        self.assertExcelEqual()
