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

        filename = 'table14.xlsx'

        test_dir = 'xlsxwriter/test/comparison/'
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + 'xlsx_files/' + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with tables."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        format1 = workbook.add_format({'num_format': '0.00;[Red]0.00', 'dxf_index': 2})
        format2 = workbook.add_format({'num_format': '0.00_ ;\-0.00\ ', 'dxf_index': 1})
        format3 = workbook.add_format({'num_format': '0.00_ ;[Red]\-0.00\ ', 'dxf_index': 0})

        data = [
            ['Foo', 1234, 2000, 4321],
            ['Bar', 1256, 4000, 4320],
            ['Baz', 2234, 3000, 4332],
            ['Bop', 1324, 1000, 4333],
        ]

        worksheet.set_column('C:F', 10.288)

        worksheet.add_table('C2:F6', {'data': data,
                                      'columns': [{},
                                                  {'format': format1},
                                                  {'format': format2},
                                                  {'format': format3},
                                                  ]})

        workbook.close()

        self.assertExcelEqual()
