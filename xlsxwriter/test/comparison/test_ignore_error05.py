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

        self.set_filename('ignore_error05.xlsx')

        self.ignore_files = ['xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels']

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.write_string('A1', '123')
        worksheet.write_formula('A2', '=1/0', None, '#DIV/0!')

        worksheet.ignore_errors({'number_stored_as_text': 'A1', 'eval_error': 'A2'})

        workbook.close()

        self.assertExcelEqual()
