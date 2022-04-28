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

        self.set_filename('simple07.xlsx')

    def test_write_nan(self):
        """Test write with NAN/INF. Issue #30"""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        worksheet.write_string(0, 0, 'Foo')
        worksheet.write_number(1, 0, 123)
        worksheet.write_string(2, 0, 'NAN')
        worksheet.write_string(3, 0, 'nan')
        worksheet.write_string(4, 0, 'INF')
        worksheet.write_string(5, 0, 'infinity')

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_in_memory(self):
        """Test write with NAN/INF. Issue #30"""

        workbook = Workbook(self.got_filename, {'in_memory': True})
        worksheet = workbook.add_worksheet()

        worksheet.write_string(0, 0, 'Foo')
        worksheet.write_number(1, 0, 123)
        worksheet.write_string(2, 0, 'NAN')
        worksheet.write_string(3, 0, 'nan')
        worksheet.write_string(4, 0, 'INF')
        worksheet.write_string(5, 0, 'infinity')

        workbook.close()

        self.assertExcelEqual()
