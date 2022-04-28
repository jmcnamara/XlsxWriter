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

        self.set_filename('simple02.xlsx')

    def test_create_file(self):
        """Test the creation of a simple workbook."""

        workbook = Workbook(self.got_filename)

        worksheet1 = workbook.add_worksheet()
        workbook.add_worksheet('Data Sheet')
        worksheet3 = workbook.add_worksheet()

        bold = workbook.add_format({'bold': 1})

        worksheet1.write_string(0, 0, 'Foo')
        worksheet1.write_number(1, 0, 123)

        worksheet3.write_string(1, 1, 'Foo')
        worksheet3.write_string(2, 1, 'Bar', bold)
        worksheet3.write_number(3, 2, 234)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_A1(self):
        """Test the creation of a simple workbook with A1 notation."""

        workbook = Workbook(self.got_filename)

        worksheet1 = workbook.add_worksheet()
        worksheet2 = workbook.add_worksheet('Data Sheet')
        worksheet3 = workbook.add_worksheet()

        bold = workbook.add_format({'bold': 1})

        worksheet1.write('A1', 'Foo')
        worksheet1.write('A2', 123)

        worksheet3.write('B2', 'Foo')
        worksheet3.write('B3', 'Bar', bold)
        worksheet3.write('C4', 234)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_in_memory(self):
        """Test the creation of a simple workbook."""

        workbook = Workbook(self.got_filename, {'in_memory': True})

        worksheet1 = workbook.add_worksheet()
        workbook.add_worksheet('Data Sheet')
        worksheet3 = workbook.add_worksheet()

        bold = workbook.add_format({'bold': 1})

        worksheet1.write_string(0, 0, 'Foo')
        worksheet1.write_number(1, 0, 123)

        worksheet3.write_string(1, 1, 'Foo')
        worksheet3.write_string(2, 1, 'Bar', bold)
        worksheet3.write_number(3, 2, 234)

        workbook.close()

        self.assertExcelEqual()
