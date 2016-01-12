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

        filename = 'excel2003_style07.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.image_dir = test_dir + 'images/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {'xl/drawings/drawing1.xml':
                                ['<xdr:cNvPr', '<a:picLocks',
                                 '<a:srcRect/>', '<xdr:spPr',
                                 '<a:noFill/>']}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename, {'excel2003_style': True})

        worksheet = workbook.add_worksheet()

        worksheet.insert_image('B3', self.image_dir + 'yellow.jpg',
                               {'x_offset': 4, 'y_offset': 3})

        workbook.close()

        self.assertExcelEqual()
