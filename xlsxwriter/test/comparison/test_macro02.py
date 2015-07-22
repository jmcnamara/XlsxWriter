###############################################################################
#
# Tests for XlsxWriter.
#
# This test verifies that any signature files and custom-ui files are included
# in the xlsx ouput
from io import BytesIO

from ..excel_comparsion_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.maxDiff = None

        filename = 'macro02.xlsm'

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
        workbook.add_vba_project(self.vba_dir + 'vbaProject03.bin',
                                 signature=self.vba_dir + 'vbaProjectSignature03.bin')
        workbook.add_custom_ui(self.vba_dir + 'customUI-01.xml', version=2006)
        workbook.add_custom_ui(self.vba_dir + 'customUI14-01.xml', version=2007)
        worksheet.write('A1', 'Test')
        workbook.close()

        self.assertExcelEqual()

    def test_create_file_in_memory(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename, {'in_memory': True})
        worksheet = workbook.add_worksheet()
        workbook.add_vba_project(self.vba_dir + 'vbaProject03.bin',
                                 signature=self.vba_dir + 'vbaProjectSignature03.bin')
        workbook.add_custom_ui(self.vba_dir + 'customUI-01.xml', version=2006)
        workbook.add_custom_ui(self.vba_dir + 'customUI14-01.xml', version=2007)

        worksheet.write('A1', 'Test')

        workbook.close()
        self.assertExcelEqual()

    def test_create_file_bytes_io(self):
        """Test the creation of a simple XlsxWriter file."""
        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        vba_file = open(self.vba_dir + 'vbaProject03.bin', 'rb')
        vba_signature_file = open(self.vba_dir + 'vbaProjectSignature03.bin', 'rb')
        vba_data = BytesIO(vba_file.read())
        vba_signature_data = BytesIO(vba_signature_file.read())
        vba_file.close()
        vba_signature_file.close()

        workbook.add_vba_project(vba_data, True, vba_signature_data, True)
        workbook.add_custom_ui(self.vba_dir + 'customUI-01.xml', version=2006)
        workbook.add_custom_ui(self.vba_dir + 'customUI14-01.xml', version=2007)

        worksheet.write('A1', 'Test')

        workbook.close()
        self.assertExcelEqual()

    def test_create_file_bytes_io_in_memory(self):
        """Test the creation of a simple XlsxWriter file."""
        workbook = Workbook(self.got_filename, {'in_memory': True})
        worksheet = workbook.add_worksheet()

        vba_file = open(self.vba_dir + 'vbaProject03.bin', 'rb')
        vba_signature_file = open(self.vba_dir + 'vbaProjectSignature03.bin', 'rb')
        vba_data = BytesIO(vba_file.read())
        vba_signature_data = BytesIO(vba_signature_file.read())
        vba_file.close()
        vba_signature_file.close()

        workbook.add_vba_project(vba_data, True, vba_signature_data, True)
        workbook.add_custom_ui(self.vba_dir + 'customUI-01.xml', version=2006)
        workbook.add_custom_ui(self.vba_dir + 'customUI14-01.xml', version=2007)

        worksheet.write('A1', 'Test')

        workbook.close()
        self.assertExcelEqual()
