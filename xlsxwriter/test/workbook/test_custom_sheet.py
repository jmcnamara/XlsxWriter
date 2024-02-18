###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest

from xlsxwriter.chartsheet import Chartsheet
from xlsxwriter.worksheet import Worksheet
from ...workbook import Workbook


class MyWorksheet(Worksheet):
    pass


class MyChartsheet(Chartsheet):
    pass


class MyWorkbook(Workbook):
    chartsheet_class = MyChartsheet
    worksheet_class = MyWorksheet


class TestCustomSheet(unittest.TestCase):
    """
    Test the Workbook _check_sheetname() method.

    """

    def setUp(self):
        self.workbook = Workbook()

    def tearDown(self):
        self.workbook.fileclosed = 1

    def test_check_chartsheet(self):
        """Test the _check_sheetname() method"""
        sheet = self.workbook.add_chartsheet()
        assert isinstance(sheet, Chartsheet)

        sheet = self.workbook.add_chartsheet(chartsheet_class=MyChartsheet)
        assert isinstance(sheet, MyChartsheet)

    def test_check_worksheet(self):
        """Test the _check_sheetname() method"""
        sheet = self.workbook.add_worksheet()
        assert isinstance(sheet, Worksheet)

        sheet = self.workbook.add_worksheet(worksheet_class=MyWorksheet)
        assert isinstance(sheet, MyWorksheet)


class TestCustomWorkBook(unittest.TestCase):
    """
    Test the Workbook _check_sheetname() method.

    """

    def setUp(self):
        self.workbook = MyWorkbook()

    def tearDown(self):
        self.workbook.fileclosed = 1

    def test_check_chartsheet(self):
        """Test the _check_sheetname() method"""
        sheet = self.workbook.add_chartsheet()
        assert isinstance(sheet, MyChartsheet)

    def test_check_worksheet(self):
        """Test the _check_sheetname() method"""
        sheet = self.workbook.add_worksheet()
        assert isinstance(sheet, MyWorksheet)
