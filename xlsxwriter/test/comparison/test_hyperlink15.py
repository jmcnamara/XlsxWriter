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
        self.set_filename("hyperlink15.xlsx")

    def test_create_file(self):
        """
        Test the creation of a simple XlsxWriter file with hyperlinks. This
        example doesn't have any link formatting and tests the relationship
        linkage code.

        """

        workbook = Workbook(self.got_filename)

        # Turn off default URL format for testing.
        workbook.default_url_format = None

        worksheet = workbook.add_worksheet()

        worksheet.write_url("B2", "external:subdir/blank.xlsx")

        workbook.close()

        self.assertExcelEqual()
