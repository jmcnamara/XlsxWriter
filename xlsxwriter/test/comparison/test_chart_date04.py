###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

from datetime import date
from ..excel_comparsion_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.maxDiff = None

        filename = 'chart_date04.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {'xl/charts/chart1.xml': ['<c:formatCode']}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({'type': 'line'})
        date_format = workbook.add_format({'num_format': 14})

        chart.axis_ids = [51761152, 51762688]

        worksheet.set_column('A:A', 12)

        dates = [date(2013, 1, 1),
                 date(2013, 1, 2),
                 date(2013, 1, 3),
                 date(2013, 1, 4),
                 date(2013, 1, 5),
                 date(2013, 1, 6),
                 date(2013, 1, 7),
                 date(2013, 1, 8),
                 date(2013, 1, 9),
                 date(2013, 1, 10)]

        values = [10, 30, 20, 40, 20, 60, 50, 40, 30, 30]

        worksheet.write_column('A1', dates, date_format)
        worksheet.write_column('B1', values)

        chart.add_series({
            'categories': '=Sheet1!$A$1:$A$10',
            'values': '=Sheet1!$B$1:$B$10',
        })

        chart.set_x_axis({
            'date_axis': True,
            'minor_unit': 1,
            'major_unit': 1,
            'minor_unit_type': 'months',
            'major_unit_type': 'years',
            'num_format': 'dd/mm/yyyy',
            'num_format_linked': True,
        })

        worksheet.insert_chart('E9', chart)

        workbook.close()

        self.assertExcelEqual()
