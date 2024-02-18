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
        self.set_filename("autofit12.xlsx")

        self.ignore_files = [
            "xl/calcChain.xml",
            "[Content_Types].xml",
            "xl/_rels/workbook.xml.rels",
        ]

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.write_array_formula(0, 0, 2, 0, "{=SUM(B1:C1*B2:C2)}", None, 1000)

        worksheet.write(0, 1, 20)
        worksheet.write(1, 1, 30)
        worksheet.write(2, 1, 40)

        worksheet.write(0, 2, 10)
        worksheet.write(1, 2, 40)
        worksheet.write(2, 2, 20)

        worksheet.autofit()

        worksheet.write(1, 0, 1000)
        worksheet.write(2, 0, 1000)

        workbook.close()

        self.assertExcelEqual()
