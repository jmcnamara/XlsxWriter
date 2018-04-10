###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

from __future__ import with_statement
from ..excel_comparsion_test import ExcelComparisonTest
from datetime import date
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.maxDiff = None

        filename = 'simple01.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple workbook."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        worksheet.write_string(0, 0, 'Hello')
        worksheet.write_number(1, 0, 123)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_A1(self):
        """Test the creation of a simple workbook with A1 notation."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        worksheet.write_string('A1', 'Hello')
        worksheet.write_number('A2', 123)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_write(self):
        """Test the creation of a simple workbook using write()."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        worksheet.write(0, 0, 'Hello')
        worksheet.write(1, 0, 123)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_with_statement(self):
        """Test the creation of a simple workbook using `with` statement."""

        with Workbook(self.got_filename) as workbook:
            worksheet = workbook.add_worksheet()

            worksheet.write(0, 0, 'Hello')
            worksheet.write(1, 0, 123)

        self.assertExcelEqual()

    def test_create_file_write_A1(self):
        """Test the creation of a simple workbook using write() with A1."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        worksheet.write('A1', 'Hello')
        worksheet.write('A2', 123)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_kwargs(self):
        """Test the creation of a simple workbook with keyword args."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        worksheet.write_string(row=0, col=0, string='Hello')
        worksheet.write_number(row=1, col=0, number=123)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_write_date_default(self):
        """Test writing a datetime without a format. Issue #33"""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        worksheet.write('A1', 'Hello')
        worksheet.write('A2', date(1900, 5, 2))

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_in_memory(self):
        """Test the creation of a simple workbook."""

        workbook = Workbook(self.got_filename, {'in_memory': True})
        worksheet = workbook.add_worksheet()

        worksheet.write_string(0, 0, 'Hello')
        worksheet.write_number(1, 0, 123)

        workbook.close()

        self.assertExcelEqual()
