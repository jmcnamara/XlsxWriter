###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparsion_test import ExcelComparisonTest
import sys
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.maxDiff = None

        filename = 'shared_strings03.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        if sys.version_info[0] == 2:
            non_char1 = unichr(0xFFFE)
            non_char2 = unichr(0xFFFF)
        else:
            non_char1 = "\uFFFE"
            non_char2 = "\uFFFF"

        worksheet.write(0, 0, non_char1)
        worksheet.write(1, 0, non_char2)

        workbook.close()

        self.assertExcelEqual()
