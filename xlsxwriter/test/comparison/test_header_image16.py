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

        self.set_filename('header_image16.xlsx')
        self.ignore_elements = {'xl/worksheets/sheet1.xml': ['<pageMargins', '<pageSetup'],
                                'xl/worksheets/sheet2.xml': ['<pageMargins', '<pageSetup']}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with image(s)."""

        workbook = Workbook(self.got_filename)

        worksheet1 = workbook.add_worksheet()
        worksheet2 = workbook.add_worksheet()

        worksheet1.set_header('&L&G', {'image_left': self.image_dir + 'red.jpg'})
        worksheet2.set_header('&L&G', {'image_left': self.image_dir + 'red.jpg'})

        worksheet1.set_footer('&R&G', {'image_right': self.image_dir + 'red.jpg'})
        worksheet2.set_footer('&R&G', {'image_right': self.image_dir + 'red.jpg'})

        workbook.close()

        self.assertExcelEqual()
