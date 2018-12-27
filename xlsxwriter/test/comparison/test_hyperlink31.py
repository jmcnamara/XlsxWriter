###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparsion_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):

        self.set_filename('hyperlink31.xlsx')

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with hyperlinks."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        format1 = workbook.add_format({'bold': True})

        worksheet.write('A1', 'Test', format1)
        worksheet.write('A3', 'http://www.python.org/')

        workbook.close()

        self.assertExcelEqual()
