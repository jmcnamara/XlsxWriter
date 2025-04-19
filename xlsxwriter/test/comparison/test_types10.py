###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

import uuid

from xlsxwriter.workbook import Workbook

from ..excel_comparison_test import ExcelComparisonTest


def write_uuid(worksheet, row, col, token, cell_format=None):
    return worksheet.write_string(row, col, str(token), cell_format)


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("types10.xlsx")

    def test_write_user_type(self):
        """Test writing numbers as text."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        worksheet.add_write_handler(uuid.UUID, write_uuid)
        my_uuid = uuid.uuid3(uuid.NAMESPACE_DNS, "python.org")

        worksheet.write("A1", my_uuid)

        workbook.close()

        self.assertExcelEqual()
