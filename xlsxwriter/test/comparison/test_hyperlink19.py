###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

from xlsxwriter.sharedstrings import SharedStringTable
from xlsxwriter.workbook import Workbook

from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("hyperlink19.xlsx")

        self.ignore_files = [
            "xl/calcChain.xml",
            "[Content_Types].xml",
            "xl/_rels/workbook.xml.rels",
        ]

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with hyperlinks."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.write_url("A1", "http://www.perl.com/")

        # Maintain the link but overwrite string with a formula.
        worksheet.write_formula("A1", "=1+1", None, 2)

        # Reset the SST for testing.
        workbook.str_table = SharedStringTable()

        workbook.close()

        self.assertExcelEqual()
