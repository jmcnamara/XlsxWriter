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
        self.set_filename("protect08.xlsx")

    def test_create_file(self):
        """Test the a simple XlsxWriter file with worksheet protection."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        options = {
            "objects": True,
            "scenarios": True,
            "format_cells": True,
            "format_columns": True,
            "format_rows": True,
            "insert_columns": True,
            "insert_rows": True,
            "insert_hyperlinks": True,
            "delete_columns": True,
            "delete_rows": True,
            "select_locked_cells": False,
            "sort": True,
            "autofilter": True,
            "pivot_tables": True,
            "select_unlocked_cells": False,
        }

        worksheet.protect("", options)

        workbook.close()

        self.assertExcelEqual()
