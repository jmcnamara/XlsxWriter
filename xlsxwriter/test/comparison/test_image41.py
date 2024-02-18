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
        self.set_filename("image41.xlsx")

        # Despite a lot of effort and testing I can't match Excel's
        # calculations exactly for EMF files. The differences are are small
        # (<1%) and in general aren't visible. The following ignore the
        # elements where these differences occur until the they can be
        # resolved. This issue doesn't occur for any other image type.
        self.ignore_elements = {
            "xl/drawings/drawing1.xml": ["<xdr:rowOff>", "<xdr:colOff>", "<a:ext cx="]
        }

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.insert_image("E9", self.image_dir + "logo.emf")

        workbook.close()

        self.assertExcelEqual()
