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

        filename = 'chart_scatter15.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({'type': 'scatter'})

        chart.axis_ids = [58843520, 58845440]

        data = [
            ['X', 1, 3],
            ['Y', 10, 30],
        ]

        worksheet.write_column('A1', data[0])
        worksheet.write_column('B1', data[1])

        chart.add_series({
            'categories': '=Sheet1!$A$2:$A$3',
            'values': '=Sheet1!$B$2:$B$3',
        })

        chart.set_x_axis({'name': '=Sheet1!$A$1',
                          'name_font': {'italic': 1, 'baseline': -1}})

        chart.set_y_axis({'name': '=Sheet1!$B$1'})

        worksheet.insert_chart('E9', chart)

        workbook.close()

        self.assertExcelEqual()
