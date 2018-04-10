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

        filename = 'cond_format18.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with conditionalformatting."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.write('A1', 1)
        worksheet.write('A2', 2)
        worksheet.write('A3', 3)
        worksheet.write('A4', 4)
        worksheet.write('A5', 5)
        worksheet.write('A6', 6)
        worksheet.write('A7', 7)
        worksheet.write('A8', 8)
        worksheet.write('A9', 9)

        worksheet.write('A12', 75)

        worksheet.conditional_format('A1',
                                     {'type': 'icon_set',
                                      'icon_style': '3_arrows',
                                      'reverse_icons': True})

        worksheet.conditional_format('A2',
                                     {'type': 'icon_set',
                                      'icon_style': '3_flags',
                                      'icons_only': True})

        worksheet.conditional_format('A3',
                                     {'type': 'icon_set',
                                      'icon_style': '3_traffic_lights_rimmed',
                                      'icons_only': True,
                                      'reverse_icons': True})

        worksheet.conditional_format('A4',
                                     {'type': 'icon_set',
                                      'icon_style': '3_symbols_circled',
                                      'icons': [{'value': 80},
                                                {'value': 20}]})

        worksheet.conditional_format('A5',
                                     {'type': 'icon_set',
                                      'icon_style': '4_arrows',
                                      'icons': [{'criteria': '>'},
                                                {'criteria': '>'},
                                                {'criteria': '>'}]})

        worksheet.conditional_format('A6',
                                     {'type': 'icon_set',
                                      'icon_style': '4_red_to_black',
                                      'icons': [{'criteria': '>=', 'type': 'number', 'value': 90},
                                                {'criteria': '<', 'type': 'percentile', 'value': 50},
                                                {'criteria': '<=', 'type': 'percent', 'value': 25}]})

        worksheet.conditional_format('A7',
                                     {'type': 'icon_set',
                                      'icon_style': '4_traffic_lights',
                                      'icons': [{'value': '=$A$12'}]})

        worksheet.conditional_format('A8',
                                     {'type': 'icon_set',
                                      'icon_style': '5_arrows_gray',
                                      'icons': [{'type': 'formula', 'value': '=$A$12'}]})

        worksheet.conditional_format('A9',
                                     {'type': 'icon_set',
                                      'icon_style': '5_quarters',
                                      'icons': [{'value': 70},
                                                {'value': 50},
                                                {'value': 30},
                                                {'value': 10}],
                                      'reverse_icons': True})

        workbook.close()

        self.assertExcelEqual()
