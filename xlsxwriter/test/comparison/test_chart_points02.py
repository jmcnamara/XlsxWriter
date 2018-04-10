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

        filename = 'chart_points02.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of an XlsxWriter file with point formatting."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({'type': 'pie'})

        data = [2, 5, 4, 1, 7, 4]

        worksheet.write_column('A1', data)

        chart.add_series({
            'values': '=Sheet1!$A$1:$A$6',
            'points': [
                None, {'border': {'color': 'red', 'dash_type': 'square_dot'}},
                None, {'fill': {'color': 'yellow'}}
            ],
        })

        worksheet.insert_chart('E9', chart)

        workbook.close()

        self.assertExcelEqual()
