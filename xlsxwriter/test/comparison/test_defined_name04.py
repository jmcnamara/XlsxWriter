###############################################################################
# _*_ coding: utf-8
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#
from __future__ import unicode_literals
from ..excel_comparsion_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.maxDiff = None

        filename = 'defined_name04.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with defined names."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        # Test for valid Excel defined names.
        workbook.define_name('\\__', '=Sheet1!$A$1')
        workbook.define_name('a3f6', '=Sheet1!$A$2')
        workbook.define_name('afoo.bar', '=Sheet1!$A$3')
        workbook.define_name('étude', '=Sheet1!$A$4')
        workbook.define_name('eésumé', '=Sheet1!$A$5')
        workbook.define_name('b', '=Sheet1!$A$6')

        # The following aren't valid Excel names and shouldn't be written to
        # the output file. We ignore the warnings raised in define_name() and
        # instead check that the output file only contains the valid names.
        import warnings
        warnings.filterwarnings('ignore')

        workbook.define_name('.abc', '=Sheet1!$B$1')
        workbook.define_name('GFG$', '=Sheet1!$B$1')
        workbook.define_name('A1', '=Sheet1!$B$1')
        workbook.define_name('XFD1048576', '=Sheet1!$B$1')
        workbook.define_name('1A', '=Sheet1!$B$1')
        workbook.define_name('A A', '=Sheet1!$B$1')
        workbook.define_name('c', '=Sheet1!$B$1')
        workbook.define_name('r', '=Sheet1!$B$1')
        workbook.define_name('C', '=Sheet1!$B$1')
        workbook.define_name('R', '=Sheet1!$B$1')
        workbook.define_name('R1', '=Sheet1!$B$1')
        workbook.define_name('C1', '=Sheet1!$B$1')
        workbook.define_name('R1C1', '=Sheet1!$B$1')
        workbook.define_name('R13C99', '=Sheet1!$B$1')

        warnings.resetwarnings()

        workbook.close()

        self.assertExcelEqual()
