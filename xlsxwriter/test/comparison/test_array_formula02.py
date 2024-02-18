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
        self.set_filename("array_formula02.xlsx")

        self.ignore_files = [
            "xl/calcChain.xml",
            "[Content_Types].xml",
            "xl/_rels/workbook.xml.rels",
        ]

    def test_create_file(self):
        """Test the creation of an XlsxWriter file with an array formula."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        xf_format = workbook.add_format({"bold": 1})

        worksheet.write("B1", 0)
        worksheet.write("B2", 0)
        worksheet.write("B3", 0)
        worksheet.write("C1", 0)
        worksheet.write("C2", 0)
        worksheet.write("C3", 0)

        worksheet.write_array_formula(0, 0, 2, 0, "{=SUM(B1:C1*B2:C2)}", xf_format, 0)

        workbook.close()

        self.assertExcelEqual()
