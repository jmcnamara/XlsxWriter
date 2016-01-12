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

        filename = 'chart_bar11.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        # Turn off default URL format for testing.
        workbook.default_url_format = None

        worksheet = workbook.add_worksheet()
        chart1 = workbook.add_chart({'type': 'bar'})
        chart2 = workbook.add_chart({'type': 'bar'})
        chart3 = workbook.add_chart({'type': 'bar'})

        chart1.axis_ids = [40274944, 40294272]
        chart2.axis_ids = [62355328, 62356864]
        chart3.axis_ids = [79538816, 65422464]

        data = [
            [1, 2, 3, 4, 5],
            [2, 4, 6, 8, 10],
            [3, 6, 9, 12, 15],
        ]

        worksheet.write_column('A1', data[0])
        worksheet.write_column('B1', data[1])
        worksheet.write_column('C1', data[2])

        worksheet.write('A7', 'http://www.perl.com/')
        worksheet.write('A8', 'http://www.perl.org/')
        worksheet.write('A9', 'http://www.perl.net/')

        chart1.add_series({'values': '=Sheet1!$A$1:$A$5'})
        chart1.add_series({'values': '=Sheet1!$B$1:$B$5'})
        chart1.add_series({'values': '=Sheet1!$C$1:$C$5'})

        chart2.add_series({'values': '=Sheet1!$A$1:$A$5'})
        chart2.add_series({'values': '=Sheet1!$B$1:$B$5'})

        chart3.add_series({'values': '=Sheet1!$A$1:$A$5'})

        worksheet.insert_chart('E9', chart1)
        worksheet.insert_chart('D25', chart2)
        worksheet.insert_chart('L32', chart3)

        workbook.close()

        self.assertExcelEqual()
