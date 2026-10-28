###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import warnings

from xlsxwriter.workbook import Workbook

from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("table01.xlsx")

        # Ignore increased shared string count.
        self.ignore_files = ["xl/sharedStrings.xml"]

    def test_create_file(self):
        """Test tables with headers that duplicate the default headers."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        worksheet.set_column("C:F", 10.288)

        worksheet.add_table("C3:F13")

        # The following tables should be ignored since the user defined
        # headers duplicate the default "ColumnN" headers. See issue #1215.
        warnings.filterwarnings("ignore")

        worksheet.add_table("G3:H4", {"columns": [{"header": "Column2"}, {}]})
        worksheet.add_table("I3:J4", {"columns": [{}, {"header": "Column1"}]})
        worksheet.add_table("K3:L4", {"columns": [{"header": "column2"}]})

        workbook.close()

        self.assertExcelEqual()
