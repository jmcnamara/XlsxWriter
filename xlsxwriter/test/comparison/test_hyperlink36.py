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

        self.set_filename('hyperlink36.xlsx')

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({'type': 'pie'})

        worksheet.write('A1', 1)
        worksheet.write('A2', 2)

        worksheet.insert_image('E9',
                               self.image_dir + 'red.png',
                               {'url': 'https://github.com/jmcnamara'})

        chart.add_series({'values': '=Sheet1!$A$1:$A$2'})

        worksheet.insert_chart('E12', chart)

        workbook.close()

        self.assertExcelEqual()
