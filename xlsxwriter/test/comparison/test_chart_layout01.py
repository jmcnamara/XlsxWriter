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

        filename = 'chart_layout01.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of an XlsxWriter file with user defined layout."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart = workbook.add_chart({'type': 'column'})

        chart.axis_ids = [69198592, 69200128]

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

        chart.set_plotarea({
            'layout': {
                'x': 0.13171062992125984,
                'y': 0.26436351706036748,
                'width': 0.73970734908136482,
                'height': 0.5713732137649461,
            }
        })

        worksheet.insert_chart('E9', chart)

        workbook.close()

        self.assertExcelEqual()
