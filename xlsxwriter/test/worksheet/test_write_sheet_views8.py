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

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_write_sheet_views1(self):
        """Test the _write_sheet_views() method with split panes + selection"""

        self.worksheet.select()

        self.worksheet.set_selection('A2')
        self.worksheet.split_panes(15, 0)

        self.worksheet._write_sheet_views()

        exp = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="600" topLeftCell="A2" activePane="bottomLeft"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>'
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_views2(self):
        """Test the _write_sheet_views() method with split panes + selection"""

        self.worksheet.select()

        self.worksheet.set_selection('B1')
        self.worksheet.split_panes(0, 8.43)

        self.worksheet._write_sheet_views()

        exp = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1350" topLeftCell="B1" activePane="topRight"/><selection pane="topRight" activeCell="B1" sqref="B1"/></sheetView></sheetViews>'
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_views3(self):
        """Test the _write_sheet_views() method with split panes + selection"""

        self.worksheet.select()

        self.worksheet.set_selection('G4')
        self.worksheet.split_panes(45, 54.14)

        self.worksheet._write_sheet_views()

        exp = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="6150" ySplit="1200" topLeftCell="G4" activePane="bottomRight"/><selection pane="topRight" activeCell="G1" sqref="G1"/><selection pane="bottomLeft" activeCell="A4" sqref="A4"/><selection pane="bottomRight" activeCell="G4" sqref="G4"/></sheetView></sheetViews>'
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_views4(self):
        """Test the _write_sheet_views() method with split panes + selection"""

        self.worksheet.select()

        self.worksheet.set_selection('I5')
        self.worksheet.split_panes(45, 54.14)

        self.worksheet._write_sheet_views()

        exp = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="6150" ySplit="1200" topLeftCell="G4" activePane="bottomRight"/><selection pane="topRight" activeCell="G1" sqref="G1"/><selection pane="bottomLeft" activeCell="A4" sqref="A4"/><selection pane="bottomRight" activeCell="I5" sqref="I5"/></sheetView></sheetViews>'
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
