###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
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
        """Test the _write_sheet_views() method with freeze panes"""

        self.worksheet.select()

        self.worksheet.freeze_panes(1, 0)

        self.worksheet._write_sheet_views()

        exp = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/></sheetView></sheetViews>'
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_views2(self):
        """Test the _write_sheet_views() method with freeze panes"""

        self.worksheet.select()

        self.worksheet.freeze_panes(0, 1)

        self.worksheet._write_sheet_views()

        exp = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/><selection pane="topRight"/></sheetView></sheetViews>'
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_views3(self):
        """Test the _write_sheet_views() method with freeze panes"""

        self.worksheet.select()

        self.worksheet.freeze_panes(1, 1)

        self.worksheet._write_sheet_views()

        exp = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/><selection pane="topRight" activeCell="B1" sqref="B1"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/><selection pane="bottomRight"/></sheetView></sheetViews>'
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_views4(self):
        """Test the _write_sheet_views() method with freeze panes"""

        self.worksheet.select()

        self.worksheet.freeze_panes("G4")

        self.worksheet._write_sheet_views()

        exp = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="6" ySplit="3" topLeftCell="G4" activePane="bottomRight" state="frozen"/><selection pane="topRight" activeCell="G1" sqref="G1"/><selection pane="bottomLeft" activeCell="A4" sqref="A4"/><selection pane="bottomRight"/></sheetView></sheetViews>'
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_views5(self):
        """Test the _write_sheet_views() method with freeze panes"""

        self.worksheet.select()

        self.worksheet.freeze_panes(3, 6, 3, 6, 1)

        self.worksheet._write_sheet_views()

        exp = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="6" ySplit="3" topLeftCell="G4" activePane="bottomRight" state="frozenSplit"/><selection pane="topRight" activeCell="G1" sqref="G1"/><selection pane="bottomLeft" activeCell="A4" sqref="A4"/><selection pane="bottomRight"/></sheetView></sheetViews>'
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
