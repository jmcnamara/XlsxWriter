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
        self.set_filename("hyperlink03.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with hyperlinks."""

        workbook = Workbook(self.got_filename)

        # Turn off default URL format for testing.
        workbook.default_url_format = None

        worksheet1 = workbook.add_worksheet()
        worksheet2 = workbook.add_worksheet()

        worksheet1.write_url("A1", "http://www.perl.org/")
        worksheet1.write_url("D4", "http://www.perl.org/")
        worksheet1.write_url("A8", "http://www.perl.org/")
        worksheet1.write_url("B6", "http://www.cpan.org/")
        worksheet1.write_url("F12", "http://www.cpan.org/")

        worksheet2.write_url("C2", "http://www.google.com/")
        worksheet2.write_url("C5", "http://www.cpan.org/")
        worksheet2.write_url("C7", "http://www.perl.org/")

        workbook.close()

        self.assertExcelEqual()
