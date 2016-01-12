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

        filename = 'chart_bar23.xlsx'

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

        chart.axis_ids = [43706240, 43727104]

        headers = ['Series 1', 'Series 2', 'Series 3']

        data = [
            ['Category 1', 'Category 2', 'Category 3', 'Category 4'],
            [4.3, 2.5, 3.5, 4.5],
            [2.4, 4.5, 1.8, 2.8],
            [2, 2, 3, 5],
        ]

        worksheet.set_column('A:D', 11)

        worksheet.write_row('B1', headers)
        worksheet.write_column('A2', data[0])
        worksheet.write_column('B2', data[1])
        worksheet.write_column('C2', data[2])
        worksheet.write_column('D2', data[3])

        chart.add_series({
            'categories': '=Sheet1!$A$2:$A$5',
            'values': '=Sheet1!$B$2:$B$5',
        })

        chart.add_series({
            'categories': '=Sheet1!$A$2:$A$5',
            'values': '=Sheet1!$C$2:$C$5',
        })

        chart.add_series({
            'categories': '=Sheet1!$A$2:$A$5',
            'values': '=Sheet1!$D$2:$D$5',
        })

        worksheet.insert_chart('E9', chart)

        workbook.close()

        self.assertExcelEqual()
