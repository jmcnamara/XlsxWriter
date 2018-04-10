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
        self.maxDiff = None

        filename = 'outline04.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = ['xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels']
        self.ignore_elements = {}

    def test_create_file(self):
        """
        Test the creation of a outlines in a XlsxWriter file. These tests are
        based on the outline programs in the examples directory.
        """

        workbook = Workbook(self.got_filename)

        worksheet4 = workbook.add_worksheet('Outline levels')

        levels = [
            "Level 1", "Level 2", "Level 3", "Level 4", "Level 5", "Level 6",
            "Level 7", "Level 6", "Level 5", "Level 4", "Level 3", "Level 2",
            "Level 1"]

        worksheet4.write_column('A1', levels)

        worksheet4.set_row(0, None, None, {'level': 1})
        worksheet4.set_row(1, None, None, {'level': 2})
        worksheet4.set_row(2, None, None, {'level': 3})
        worksheet4.set_row(3, None, None, {'level': 4})
        worksheet4.set_row(4, None, None, {'level': 5})
        worksheet4.set_row(5, None, None, {'level': 6})
        worksheet4.set_row(6, None, None, {'level': 7})
        worksheet4.set_row(7, None, None, {'level': 6})
        worksheet4.set_row(8, None, None, {'level': 5})
        worksheet4.set_row(9, None, None, {'level': 4})
        worksheet4.set_row(10, None, None, {'level': 3})
        worksheet4.set_row(11, None, None, {'level': 2})
        worksheet4.set_row(12, None, None, {'level': 1})

        workbook.close()

        self.assertExcelEqual()
