###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2022, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparison_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):

        self.set_filename('chart_blank05.xlsx')

        self.ignore_elements = {'xl/drawings/drawing1.xml': ['<xdr:ext']}

    def test_create_file(self):
        """Test the worksheet properties of an XlsxWriter chartsheet file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()
        chartsheet = workbook.add_chartsheet()

        chart = workbook.add_chart({'type': 'line'})

        chart.axis_ids = [57619968, 57621504]

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

        chart.show_blanks_as('span')

        chartsheet.set_chart(chart)

        workbook.close()

        self.assertExcelEqual()
