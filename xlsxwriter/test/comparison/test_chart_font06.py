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

        filename = 'chart_font06.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({'type': 'bar'})

        chart.axis_ids = [49407488, 53740288]

        data = [
            [1, 2, 3, 4, 5],
            [2, 4, 6, 8, 10],
            [3, 6, 9, 12, 15],

        ]

        worksheet.write_column('A1', data[0])
        worksheet.write_column('B1', data[1])
        worksheet.write_column('C1', data[2])

        chart.add_series({'values': '=Sheet1!$A$1:$A$5'})
        chart.add_series({'values': '=Sheet1!$B$1:$B$5'})
        chart.add_series({'values': '=Sheet1!$C$1:$C$5'})

        chart.set_title({
            'name': 'Title',
            'name_font': {
                'name': 'Calibri',
                'pitch_family': 34,
                'charset': 0,
                'color': 'yellow',
            },
        })

        chart.set_x_axis({
            'name': 'XXX',
            'name_font': {
                'name': 'Courier New',
                'pitch_family': 49,
                'charset': 0,
                'color': '#92D050'
            },
            'num_font': {
                'name': 'Arial',
                'pitch_family': 34,
                'charset': 0,
                'color': '#00B0F0',
            },
        })

        chart.set_y_axis({
            'name': 'YYY',
            'name_font': {
                'name': 'Century',
                'pitch_family': 18,
                'charset': 0,
                'color': 'red'
            },
            'num_font': {
                'bold': 1,
                'italic': 1,
                'underline': 1,
                'color': '#7030A0',
            },
        })

        worksheet.insert_chart('E9', chart)

        workbook.close()

        self.assertExcelEqual()
