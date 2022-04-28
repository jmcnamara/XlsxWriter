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

        self.set_filename('textbox18.xlsx')

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with textbox(s)."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.insert_textbox('E9', 'This is some text')

        worksheet.insert_textbox('E19', 'This is some text',
                                 {'align': {'vertical': 'middle'}})

        worksheet.insert_textbox('E29', 'This is some text',
                                 {'align': {'vertical': 'bottom'}})

        worksheet.insert_textbox('E39', 'This is some text',
                                 {'align': {'vertical': 'top',
                                            'horizontal': 'center'}})

        worksheet.insert_textbox('E49', 'This is some text',
                                 {'align': {'vertical': 'middle',
                                            'horizontal': 'center'}})

        worksheet.insert_textbox('E59', 'This is some text',
                                 {'align': {'vertical': 'bottom',
                                            'horizontal': 'center'}})

        workbook.close()

        self.assertExcelEqual()
