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
        """Test the _write_sheet_views() method with selection set"""

        self.worksheet.select()

        self.worksheet.set_selection('A1')

        self.worksheet._write_sheet_views()

        exp = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>'
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_views2(self):
        """Test the _write_sheet_views() method with selection set"""

        self.worksheet.select()

        self.worksheet.set_selection('A2')

        self.worksheet._write_sheet_views()

        exp = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A2" sqref="A2"/></sheetView></sheetViews>'
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_views3(self):
        """Test the _write_sheet_views() method with selection set"""

        self.worksheet.select()

        self.worksheet.set_selection('B1')

        self.worksheet._write_sheet_views()

        exp = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="B1" sqref="B1"/></sheetView></sheetViews>'
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_views4(self):
        """Test the _write_sheet_views() method with selection set"""

        self.worksheet.select()

        self.worksheet.set_selection('D3')

        self.worksheet._write_sheet_views()

        exp = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="D3" sqref="D3"/></sheetView></sheetViews>'
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_views5(self):
        """Test the _write_sheet_views() method with selection set"""

        self.worksheet.select()

        self.worksheet.set_selection('D3:F4')

        self.worksheet._write_sheet_views()

        exp = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="D3" sqref="D3:F4"/></sheetView></sheetViews>'
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_views6(self):
        """Test the _write_sheet_views() method with selection set"""

        self.worksheet.select()

        # With reversed selection direction.
        self.worksheet.set_selection('F4:D3')

        self.worksheet._write_sheet_views()

        exp = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="F4" sqref="D3:F4"/></sheetView></sheetViews>'
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_views7(self):
        """Test the _write_sheet_views() method with selection set"""

        self.worksheet.select()

        # Should be the same as 'A2'
        self.worksheet.set_selection('A2:A2')

        self.worksheet._write_sheet_views()

        exp = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A2" sqref="A2"/></sheetView></sheetViews>'
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
