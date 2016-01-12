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

        filename = 'chart_bar04.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet1 = workbook.add_worksheet()
        worksheet2 = workbook.add_worksheet()
        chart1 = workbook.add_chart({'type': 'bar'})
        chart2 = workbook.add_chart({'type': 'bar'})

        chart1.axis_ids = [64446848, 64448384]
        chart2.axis_ids = [85389696, 85391232]

        data = [
            [1, 2, 3, 4, 5],
            [2, 4, 6, 8, 10],
            [3, 6, 9, 12, 15],
        ]

        worksheet1.write_column('A1', data[0])
        worksheet1.write_column('B1', data[1])
        worksheet1.write_column('C1', data[2])

        chart1.add_series({
            'categories': '=Sheet1!$A$1:$A$5',
            'values': '=Sheet1!$B$1:$B$5',
        })

        chart1.add_series({
            'categories': '=Sheet1!$A$1:$A$5',
            'values': '=Sheet1!$C$1:$C$5',
        })

        worksheet1.insert_chart('E9', chart1)

        worksheet2.write_column('A1', data[0])
        worksheet2.write_column('B1', data[1])
        worksheet2.write_column('C1', data[2])

        chart2.add_series({
            'categories': '=Sheet2!$A$1:$A$5',
            'values': '=Sheet2!$B$1:$B$5',
        })

        chart2.add_series({
            'categories': '=Sheet2!$A$1:$A$5',
            'values': '=Sheet2!$C$1:$C$5',
        })

        worksheet2.insert_chart('E9', chart2)

        workbook.close()

        self.assertExcelEqual()
