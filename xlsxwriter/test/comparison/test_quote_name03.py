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

        filename = 'quote_name03.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        data = [
            [1, 2, 3, 4, 5],
            [2, 4, 6, 8, 10],
            [3, 6, 9, 12, 15],

        ]

        sheetnames = (
            'Sheet<1', 'Sheet>2', 'Sheet=3', 'Sheet@4',
            'Sheet^5', 'Sheet`6', 'Sheet7', 'Sheet~8'
        )

        for sheetname in sheetnames:

            worksheet = workbook.add_worksheet(sheetname)
            chart = workbook.add_chart({'type': 'pie'})

            worksheet.write_column('A1', data[0])
            worksheet.write_column('B1', data[1])
            worksheet.write_column('C1', data[2])

            chart.add_series({'values': [sheetname, 0, 0, 4, 0]})
            worksheet.insert_chart('E6', chart, {'x_offset': 26, 'y_offset': 17})

        workbook.close()

        self.assertExcelEqual()
