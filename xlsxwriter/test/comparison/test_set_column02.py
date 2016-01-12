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
        self.got_filename = test_dir + '_test_2_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_column("A:A", 0.083333333333333)
        worksheet.set_column("B:B", 0.166666666666667)
        worksheet.set_column("C:C", 0.250000000000000)
        worksheet.set_column("D:D", 0.333333333333333)
        worksheet.set_column("E:E", 0.416666666666667)
        worksheet.set_column("F:F", 0.500000000000000)
        worksheet.set_column("G:G", 0.583333333333333)
        worksheet.set_column("H:H", 0.666666666666666)
        worksheet.set_column("I:I", 0.750000000000000)
        worksheet.set_column("J:J", 0.833333333333333)
        worksheet.set_column("K:K", 0.916666666666666)
        worksheet.set_column("L:L", 1.000000000000000)
        worksheet.set_column("M:M", 1.142857142857140)
        worksheet.set_column("N:N", 1.285714285714290)
        worksheet.set_column("O:O", 1.428571428571430)
        worksheet.set_column("P:P", 1.571428571428570)
        worksheet.set_column("Q:Q", 1.714285714285710)
        worksheet.set_column("R:R", 1.857142857142860)
        worksheet.set_column("S:S", 2.000000000000000)
        worksheet.set_column("T:T", 2.142857142857140)
        worksheet.set_column("U:U", 2.285714285714290)
        worksheet.set_column("V:V", 2.428571428571430)
        worksheet.set_column("W:W", 2.571428571428570)
        worksheet.set_column("X:X", 2.714285714285710)
        worksheet.set_column("Y:Y", 2.857142857142860)
        worksheet.set_column("Z:Z", 3.000000000000000)
        worksheet.set_column("AB:AB", 8.571428571428570)
        worksheet.set_column("AC:AC", 8.711428571428570)
        worksheet.set_column("AD:AD", 8.857142857142860)
        worksheet.set_column("AE:AE", 9.000000000000000)
        worksheet.set_column("AF:AF", 9.142857142857140)
        worksheet.set_column("AG:AG", 9.285714285714290)

        workbook.close()

        self.assertExcelEqual()
