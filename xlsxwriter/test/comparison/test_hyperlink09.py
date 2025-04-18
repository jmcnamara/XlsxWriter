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
        self.set_filename("hyperlink09.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with hyperlinks."""

        workbook = Workbook(self.got_filename)

        # Turn off default URL format for testing.
        workbook.default_url_format = None

        worksheet = workbook.add_worksheet()

        worksheet.write_url("A1", r"external:..\foo.xlsx")
        worksheet.write_url("A3", r"external:..\foo.xlsx#Sheet1!A1")
        worksheet.write_url(
            "A5",
            r"external:\\VBOXSVR\share\foo.xlsx#Sheet1!B2",
            None,
            r"J:\foo.xlsx#Sheet1!B2",
        )

        workbook.close()

        self.assertExcelEqual()
