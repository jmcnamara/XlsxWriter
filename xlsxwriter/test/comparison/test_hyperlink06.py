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

        self.set_filename('hyperlink06.xlsx')

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with hyperlinks."""

        workbook = Workbook(self.got_filename)

        # Turn off default URL format for testing.
        workbook.default_url_format = None

        worksheet = workbook.add_worksheet()

        worksheet.write_url('A1', r'external:C:\Temp\foo.xlsx')
        worksheet.write_url('A3', r'external:C:\Temp\foo.xlsx#Sheet1!A1')
        worksheet.write_url('A5', r'external:C:\Temp\foo.xlsx#Sheet1!A1', None, 'External', 'Tip')

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_write(self):
        """Test the creation of a simple XlsxWriter file with hyperlinks with write()"""

        workbook = Workbook(self.got_filename)

        # Turn off default URL format for testing.
        workbook.default_url_format = None

        worksheet = workbook.add_worksheet()

        worksheet.write('A1', r'external:C:\Temp\foo.xlsx')
        worksheet.write('A3', r'external:C:\Temp\foo.xlsx#Sheet1!A1')
        worksheet.write('A5', r'external:C:\Temp\foo.xlsx#Sheet1!A1', None, 'External', 'Tip')

        workbook.close()

        self.assertExcelEqual()
