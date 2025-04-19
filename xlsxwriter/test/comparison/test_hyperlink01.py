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
        self.set_filename("hyperlink01.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with hyperlinks"""

        workbook = Workbook(self.got_filename)

        # Turn off default URL format for testing.
        workbook.default_url_format = None

        worksheet = workbook.add_worksheet()

        worksheet.write_url("A1", "http://www.perl.org/")

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_write(self):
        """Test the creation of a simple XlsxWriter file with hyperlinks with write()"""

        workbook = Workbook(self.got_filename)

        # Turn off default URL format for testing.
        workbook.default_url_format = None

        worksheet = workbook.add_worksheet()

        worksheet.write("A1", "http://www.perl.org/")

        workbook.close()

    def test_create_file_with_url_type(self):
        """Test the creation of a simple XlsxWriter using Url class"""

        workbook = Workbook(self.got_filename)

        # Turn off default URL format for testing.
        workbook.default_url_format = None

        worksheet = workbook.add_worksheet()

        url = Url("http://www.perl.org/")
        worksheet.write_url("A1", url)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_with_url_type_and_write(self):
        """Test the creation of a simple XlsxWriter using Url class"""

        workbook = Workbook(self.got_filename)

        # Turn off default URL format for testing.
        workbook.default_url_format = None

        worksheet = workbook.add_worksheet()

        url = Url("http://www.perl.org/")
        worksheet.write("A1", url)

        workbook.close()

        self.assertExcelEqual()
