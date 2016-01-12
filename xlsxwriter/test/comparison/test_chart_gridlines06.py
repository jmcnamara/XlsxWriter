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

        filename = 'chart_gridlines06.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test XlsxWriter gridlines."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({'type': 'scatter'})

        chart.axis_ids = [82812288, 46261376]

        data = [
            [1, 2, 3, 4, 5],
            [2, 4, 6, 8, 10],
            [3, 6, 9, 12, 15],

        ]

        worksheet.write_column('A1', data[0])
        worksheet.write_column('B1', data[1])
        worksheet.write_column('C1', data[2])

        chart.add_series({
            'categories': '=Sheet1!$A$1:$A$5',
            'values': '=Sheet1!$B$1:$B$5',
        })

        chart.add_series({
            'categories': '=Sheet1!$A$1:$A$5',
            'values': '=Sheet1!$C$1:$C$5',
        })

        chart.set_x_axis({
            'major_gridlines': {'visible': 1},
            'minor_gridlines': {'visible': 1},
        })

        chart.set_y_axis({
            'major_gridlines': {'visible': 1},
            'minor_gridlines': {'visible': 1},
        })

        worksheet.insert_chart('E9', chart)

        workbook.close()

        self.assertExcelEqual()
