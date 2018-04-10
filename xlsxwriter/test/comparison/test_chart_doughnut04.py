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

        filename = 'chart_doughnut04.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({'type': 'doughnut'})

        data = [
            [2, 4, 6],
            [60, 30, 10],
        ]

        worksheet.write_column('A1', data[0])
        worksheet.write_column('B1', data[1])
        chart.add_series({
            'categories': '=Sheet1!$A$1:$A$3',
            'values': '=Sheet1!$B$1:$B$3',
        })

        chart.set_rotation(30)

        worksheet.insert_chart('E9', chart)

        workbook.close()

        self.assertExcelEqual()
