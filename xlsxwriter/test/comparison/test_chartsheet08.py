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

        filename = 'chartsheet08.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = ['xl/printerSettings/printerSettings1.bin',
                             'xl/chartsheets/_rels/sheet1.xml.rels']
        self.ignore_elements = {'[Content_Types].xml': ['<Default Extension="bin"'],
                                'xl/chartsheets/sheet1.xml': ['<pageSetup', '<drawing']}

    def test_create_file(self):
        """Test the worksheet properties of an XlsxWriter chartsheet file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chartsheet = workbook.add_chartsheet()

        chart = workbook.add_chart({'type': 'bar'})

        chart.axis_ids = [46320256, 46335872]

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

        chartsheet.set_margins(left='0.70866141732283472',
                               right='0.70866141732283472',
                               top='0.74803149606299213',
                               bottom='0.74803149606299213')
        chartsheet.set_header('Page &P', '0.51181102362204722')
        chartsheet.set_footer('&A', '0.51181102362204722')
        chartsheet.set_chart(chart)

        workbook.close()

        self.assertExcelEqual()
