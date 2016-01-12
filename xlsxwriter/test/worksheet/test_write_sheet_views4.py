###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...worksheet import Worksheet


class TestWriteSheetViews(unittest.TestCase):
    """
    Test the Worksheet _write_sheet_views() method.
    Repeat of tests in test_write_sheet_views3.py with explicit topLeft cells.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_write_sheet_views1(self):
        """Test the _write_sheet_views() method with split panes"""

        self.worksheet.select()

        self.worksheet.split_panes(15, 0, 1, 0)

        self.worksheet._write_sheet_views()

        exp = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="600" topLeftCell="A2"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>'
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_views2(self):
        """Test the _write_sheet_views() method with split panes"""

        self.worksheet.select()

        self.worksheet.split_panes(30, 0, 2, 0)

        self.worksheet._write_sheet_views()

        exp = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="900" topLeftCell="A3"/><selection pane="bottomLeft" activeCell="A3" sqref="A3"/></sheetView></sheetViews>'
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_views3(self):
        """Test the _write_sheet_views() method with split panes"""

        self.worksheet.select()

        self.worksheet.split_panes(105, 0, 7, 0)

        self.worksheet._write_sheet_views()

        exp = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="2400" topLeftCell="A8"/><selection pane="bottomLeft" activeCell="A8" sqref="A8"/></sheetView></sheetViews>'
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_views4(self):
        """Test the _write_sheet_views() method with split panes"""

        self.worksheet.select()

        self.worksheet.split_panes(0, 8.43, 0, 1)

        self.worksheet._write_sheet_views()

        exp = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1350" topLeftCell="B1"/><selection pane="topRight" activeCell="B1" sqref="B1"/></sheetView></sheetViews>'
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_views5(self):
        """Test the _write_sheet_views() method with split panes"""

        self.worksheet.select()

        self.worksheet.split_panes(0, 17.57, 0, 2)

        self.worksheet._write_sheet_views()

        exp = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="2310" topLeftCell="C1"/><selection pane="topRight" activeCell="C1" sqref="C1"/></sheetView></sheetViews>'
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_views6(self):
        """Test the _write_sheet_views() method with split panes"""

        self.worksheet.select()

        self.worksheet.split_panes(0, 45, 0, 5)

        self.worksheet._write_sheet_views()

        exp = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="5190" topLeftCell="F1"/><selection pane="topRight" activeCell="F1" sqref="F1"/></sheetView></sheetViews>'
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_views7(self):
        """Test the _write_sheet_views() method with split panes"""

        self.worksheet.select()

        self.worksheet.split_panes(15, 8.43, 1, 1)

        self.worksheet._write_sheet_views()

        exp = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1350" ySplit="600" topLeftCell="B2"/><selection pane="topRight" activeCell="B1" sqref="B1"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/><selection pane="bottomRight" activeCell="B2" sqref="B2"/></sheetView></sheetViews>'
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_views8(self):
        """Test the _write_sheet_views() method with split panes"""

        self.worksheet.select()

        self.worksheet.split_panes(45, 54.14, 3, 6)

        self.worksheet._write_sheet_views()

        exp = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="6150" ySplit="1200" topLeftCell="G4"/><selection pane="topRight" activeCell="G1" sqref="G1"/><selection pane="bottomLeft" activeCell="A4" sqref="A4"/><selection pane="bottomRight" activeCell="G4" sqref="G4"/></sheetView></sheetViews>'
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
