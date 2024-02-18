###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#
import warnings
from ..excel_comparison_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("simple01.xlsx")

    def test_close_file_twice(self):
        """Test warning when closing workbook more than once."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        worksheet.write_string(0, 0, "Hello")
        worksheet.write_number(1, 0, 123)

        workbook.close()

        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            workbook.close()
            assert len(w) == 1
            assert issubclass(w[-1].category, UserWarning)

        self.assertExcelEqual()
