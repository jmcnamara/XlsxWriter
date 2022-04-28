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

        self.set_filename('chartsheet02.xlsx')

    def test_create_file(self):
        """Test the worksheet properties of an XlsxWriter chartsheet file."""

        workbook = Workbook(self.got_filename)

        worksheet1 = workbook.add_worksheet()
        chartsheet = workbook.add_chartsheet()
        worksheet2 = workbook.add_worksheet()

        chart = workbook.add_chart({'type': 'bar'})

        chart.axis_ids = [79858304, 79860096]

        data = [
            [1, 2, 3, 4, 5],
            [2, 4, 6, 8, 10],
            [3, 6, 9, 12, 15],

        ]

        worksheet1.write_column('A1', data[0])
        worksheet1.write_column('B1', data[1])
        worksheet1.write_column('C1', data[2])

        chart.add_series({'values': '=Sheet1!$A$1:$A$5'})
        chart.add_series({'values': '=Sheet1!$B$1:$B$5'})
        chart.add_series({'values': '=Sheet1!$C$1:$C$5'})

        chartsheet.set_chart(chart)
        chartsheet.activate()
        worksheet2.select()

        workbook.close()

        self.assertExcelEqual()
