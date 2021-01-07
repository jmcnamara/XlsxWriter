###############################################################################
# _*_ coding: utf-8
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2021, John McNamara, jmcnamara@cpan.org
#
from __future__ import unicode_literals
from ..excel_comparison_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):

        self.set_filename('utf8_08.xlsx')

        self.ignore_files = ['xl/printerSettings/printerSettings1.bin',
                             'xl/worksheets/_rels/sheet1.xml.rels']
        self.ignore_elements = {'[Content_Types].xml': ['<Default Extension="bin"']}

    def test_create_file(self):
        """Test the creation of an XlsxWriter file with utf-8 strings."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.write('A1', 'Foo')

        worksheet.set_header('&LCafé')
        worksheet.set_footer('&Rclé')

        worksheet.set_paper(9)

        workbook.close()

        self.assertExcelEqual()
