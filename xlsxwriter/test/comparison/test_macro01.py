###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparsion_test import ExcelComparisonTest
from ...workbook import Workbook
from ...compatibility import BytesIO


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.maxDiff = None

        filename = 'macro01.xlsm'

        test_dir = 'xlsxwriter/test/comparison/'
        self.vba_dir = test_dir + 'xlsx_files/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        workbook.add_vba_project(self.vba_dir + 'vbaProject01.bin')

        worksheet.write('A1', 123)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_in_memory(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename, {'in_memory': True})

        worksheet = workbook.add_worksheet()

        workbook.add_vba_project(self.vba_dir + 'vbaProject01.bin')

        worksheet.write('A1', 123)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_bytes_io(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        vba_file = open(self.vba_dir + 'vbaProject01.bin', 'rb')
        vba_data = BytesIO(vba_file.read())
        vba_file.close()

        workbook.add_vba_project(vba_data, True)

        worksheet.write('A1', 123)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_bytes_io_in_memory(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename, {'in_memory': True})

        worksheet = workbook.add_worksheet()

        vba_file = open(self.vba_dir + 'vbaProject01.bin', 'rb')
        vba_data = BytesIO(vba_file.read())
        vba_file.close()

        workbook.add_vba_project(vba_data, True)

        worksheet.write('A1', 123)

        workbook.close()

        self.assertExcelEqual()
