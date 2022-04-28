###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2022, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparison_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):

        self.set_filename('optimize07.xlsx')

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename, {'constant_memory': True, 'in_memory': False})
        worksheet = workbook.add_worksheet()

        strings = [
            "_",
            "_x",
            "_x0",
            "_x00",
            "_x000",
            "_x0000",
            "_x0000_",
            "_x005F_",
            "_x000G_",
            "_X0000_",
            "_x000a_",
            "_x000A_",
            "_x0000__x0000_",
            "__x0000__",
        ]

        worksheet.write_column(0, 0, strings)

        workbook.close()

        self.assertExcelEqual()
