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

        self.set_filename("embed_image13.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename)

        worksheet1 = workbook.add_worksheet()

        worksheet1.embed_image(0, 0, self.image_dir + "red.png")
        worksheet1.embed_image(2, 0, self.image_dir + "blue.png")
        worksheet1.embed_image(4, 0, self.image_dir + "yellow.png")

        worksheet2 = workbook.add_worksheet()

        worksheet2.embed_image(0, 0, self.image_dir + "yellow.png")
        worksheet2.embed_image(2, 0, self.image_dir + "red.png")
        worksheet2.embed_image(4, 0, self.image_dir + "blue.png")

        worksheet3 = workbook.add_worksheet()

        worksheet3.embed_image(0, 0, self.image_dir + "blue.png")
        worksheet3.embed_image(2, 0, self.image_dir + "yellow.png")
        worksheet3.embed_image(4, 0, self.image_dir + "red.png")

        workbook.close()

        self.assertExcelEqual()
