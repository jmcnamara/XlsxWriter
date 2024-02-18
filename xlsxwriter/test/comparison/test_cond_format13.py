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
        self.set_filename("cond_format04.xlsx")

    def test_create_file(self):
        """Test the creation of an XlsxWriter file with conditional formatting."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        format1 = workbook.add_format({"num_format": 2, "dxf_index": 1})
        format2 = workbook.add_format({"num_format": "0.000", "dxf_index": 0})

        worksheet.write("A1", 10)
        worksheet.write("A2", 20)
        worksheet.write("A3", 30)
        worksheet.write("A4", 40)

        options = {
            "type": "cell",
            "format": format1,
            "criteria": ">",
            "value": 2,
        }

        worksheet.conditional_format("A1", options)

        options["criteria"] = "<"
        options["value"] = 8
        options["format"] = format2

        worksheet.conditional_format("A2", options)

        workbook.close()

        self.assertExcelEqual()
