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

        filename = 'chart_pattern09.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({'type': 'column'})

        chart.axis_ids = [115359744, 115361280]

        data = [
            [2, 2, 2],
            [2, 2, 2],
            [2, 2, 2],
            [2, 2, 2],
            [2, 2, 2],
            [2, 2, 2],
            [2, 2, 2],
            [2, 2, 2],
        ]

        worksheet.write_column('A1', data[0])
        worksheet.write_column('B1', data[1])
        worksheet.write_column('C1', data[2])
        worksheet.write_column('D1', data[3])
        worksheet.write_column('E1', data[4])
        worksheet.write_column('F1', data[5])
        worksheet.write_column('G1', data[6])
        worksheet.write_column('H1', data[7])

        chart.add_series({'values': '=Sheet1!$A$1:$A$3'})
        chart.add_series({'values': '=Sheet1!$B$1:$B$3'})
        chart.add_series({'values': '=Sheet1!$C$1:$C$3'})
        chart.add_series({'values': '=Sheet1!$D$1:$D$3'})
        chart.add_series({'values': '=Sheet1!$E$1:$E$3'})
        chart.add_series({'values': '=Sheet1!$F$1:$F$3'})
        chart.add_series({'values': '=Sheet1!$G$1:$G$3'})
        chart.add_series({'values': '=Sheet1!$H$1:$H$3'})

        chart.set_chartarea({
            'pattern': {
                'pattern': 'percent_5',
                'fg_color': 'red',
                'bg_color': 'yellow',
            }
        })

        worksheet.insert_chart('E9', chart)

        workbook.close()

        self.assertExcelEqual()
