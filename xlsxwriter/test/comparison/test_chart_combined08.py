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

        filename = 'chart_combined08.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        # TODO. There are too many ignored elements here. Remove when the axis
        # writing is fixed for secondary scatter charts.
        self.ignore_elements = {'xl/charts/chart1.xml': ['<c:dispBlanksAs',
                                                         '<c:crossBetween',
                                                         '<c:tickLblPos',
                                                         '<c:auto',
                                                         '<c:valAx>',
                                                         '<c:catAx>',
                                                         '</c:valAx>',
                                                         '</c:catAx>',
                                                         '<c:crosses',
                                                         '<c:lblOffset',
                                                         '<c:lblAlgn']}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chart1 = workbook.add_chart({'type': 'column'})
        chart2 = workbook.add_chart({'type': 'scatter'})

        chart1.axis_ids = [81267328, 81297792]
        chart2.axis_ids = [81267328, 81297792]
        chart2.axis2_ids = [89510656, 84556032]

        data = [
            [2, 3, 4, 5, 6],
            [20, 25, 10, 10, 20],
            [5, 10, 15, 10, 5],
        ]

        worksheet.write_column('A1', data[0])
        worksheet.write_column('B1', data[1])
        worksheet.write_column('C1', data[2])

        chart1.add_series({
            'categories': '=Sheet1!$A$1:$A$5',
            'values': '=Sheet1!$B$1:$B$5'
        })

        chart2.add_series({
            'categories': '=Sheet1!$A$1:$A$5',
            'values': '=Sheet1!$C$1:$C$5',
            'y2_axis': 1,
        })

        chart1.combine(chart2)

        worksheet.insert_chart('E9', chart1)

        workbook.close()

        self.assertExcelEqual()
