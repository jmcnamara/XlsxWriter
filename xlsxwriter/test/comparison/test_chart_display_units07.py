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

        filename = 'chart_display_units07.xlsx'

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

        chart.axis_ids = [56159232, 61364096]

        data = [
            [10000000, 20000000, 30000000, 20000000, 10000000],
        ]

        worksheet.write_column(0, 0, data[0])

        chart.add_series({'values': '=Sheet1!$A$1:$A$5'})

        chart.set_y_axis({'display_units': 'ten_millions', 'display_units_visible': 0})

        worksheet.insert_chart('E9', chart)

        workbook.close()

        self.assertExcelEqual()
