###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparsion_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.maxDiff = None

        filename = 'set_column01.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_column("A:A", 0.08)
        worksheet.set_column("B:B", 0.17)
        worksheet.set_column("C:C", 0.25)
        worksheet.set_column("D:D", 0.33)
        worksheet.set_column("E:E", 0.42)
        worksheet.set_column("F:F", 0.5)
        worksheet.set_column("G:G", 0.58)
        worksheet.set_column("H:H", 0.67)
        worksheet.set_column("I:I", 0.75)
        worksheet.set_column("J:J", 0.83)
        worksheet.set_column("K:K", 0.92)
        worksheet.set_column("L:L", 1)
        worksheet.set_column("M:M", 1.14)
        worksheet.set_column("N:N", 1.29)
        worksheet.set_column("O:O", 1.43)
        worksheet.set_column("P:P", 1.57)
        worksheet.set_column("Q:Q", 1.71)
        worksheet.set_column("R:R", 1.86)
        worksheet.set_column("S:S", 2)
        worksheet.set_column("T:T", 2.14)
        worksheet.set_column("U:U", 2.29)
        worksheet.set_column("V:V", 2.43)
        worksheet.set_column("W:W", 2.57)
        worksheet.set_column("X:X", 2.71)
        worksheet.set_column("Y:Y", 2.86)
        worksheet.set_column("Z:Z", 3)
        worksheet.set_column("AB:AB", 8.57)
        worksheet.set_column("AC:AC", 8.71)
        worksheet.set_column("AD:AD", 8.86)
        worksheet.set_column("AE:AE", 9)
        worksheet.set_column("AF:AF", 9.14)
        worksheet.set_column("AG:AG", 9.29)

        workbook.close()

        self.assertExcelEqual()
