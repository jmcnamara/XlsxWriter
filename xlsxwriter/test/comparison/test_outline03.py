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
        self.set_filename("outline03.xlsx")

        self.ignore_files = [
            "xl/calcChain.xml",
            "[Content_Types].xml",
            "xl/_rels/workbook.xml.rels",
        ]

    def test_create_file(self):
        """
        Test the creation of a outlines in a XlsxWriter file. These tests are
        based on the outline programs in the examples directory.
        """

        workbook = Workbook(self.got_filename)

        worksheet3 = workbook.add_worksheet("Outline Columns")

        bold = workbook.add_format({"bold": 1})

        data = [
            ["Month", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Total"],
            ["North", 50, 20, 15, 25, 65, 80],
            ["South", 10, 20, 30, 50, 50, 50],
            ["East", 45, 75, 50, 15, 75, 100],
            ["West", 15, 15, 55, 35, 20, 50],
        ]

        worksheet3.set_row(0, None, bold)

        worksheet3.set_column("A:A", 10, bold)
        worksheet3.set_column("B:G", 6, None, {"level": 1})
        worksheet3.set_column("H:H", 10)

        for row, data_row in enumerate(data):
            worksheet3.write_row(row, 0, data_row)

        worksheet3.write("H2", "=SUM(B2:G2)", None, 255)
        worksheet3.write("H3", "=SUM(B3:G3)", None, 210)
        worksheet3.write("H4", "=SUM(B4:G4)", None, 360)
        worksheet3.write("H5", "=SUM(B5:G5)", None, 190)
        worksheet3.write("H6", "=SUM(H2:H5)", bold, 1015)

        workbook.close()

        self.assertExcelEqual()
