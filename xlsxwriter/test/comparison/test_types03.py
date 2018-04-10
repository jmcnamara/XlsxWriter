###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparsion_test import ExcelComparisonTest
from decimal import Decimal
from fractions import Fraction
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.maxDiff = None

        filename = 'types03.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_write_number_float(self):
        """Test writing number types."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        worksheet.write('A1', 0.5)
        worksheet.write_number('A2', 0.5)

        workbook.close()

        self.assertExcelEqual()

    def test_write_number_decimal(self):
        """Test writing number types."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        worksheet.write('A1', Decimal('0.5'))
        worksheet.write_number('A2', Decimal('0.5'))

        workbook.close()

        self.assertExcelEqual()

    def test_write_number_fraction(self):
        """Test writing number types."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        worksheet.write('A1', Fraction(1, 2))
        worksheet.write_number('A2', Fraction(2, 4))

        workbook.close()

        self.assertExcelEqual()
