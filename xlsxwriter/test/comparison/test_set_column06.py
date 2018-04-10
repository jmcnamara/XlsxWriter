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

        filename = 'set_column06.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({'type': 'line'})

        bold = workbook.add_format({'bold': 1})
        italic = workbook.add_format({'italic': 1})

        chart.axis_ids = [69197824, 69199360]

        data = [
            [1, 2, 3, 4, 5],
            [2, 4, 6, 8, 10],
            [3, 6, 9, 12, 15],
        ]

        worksheet.write('A1', 'Foo', bold)
        worksheet.write('B1', 'Bar', italic)
        worksheet.write_column('A2', data[0])
        worksheet.write_column('B2', data[1])
        worksheet.write_column('C2', data[2])

        worksheet.set_row(12, None, None, {'hidden': True})
        worksheet.set_column('F:F', None, None, {'hidden': True})

        chart.add_series({'values': '=Sheet1!$A$2:$A$6'})
        chart.add_series({'values': '=Sheet1!$B$2:$B$6'})
        chart.add_series({'values': '=Sheet1!$C$2:$C$6'})

        worksheet.insert_chart('E9', chart)

        workbook.close()

        self.assertExcelEqual()
