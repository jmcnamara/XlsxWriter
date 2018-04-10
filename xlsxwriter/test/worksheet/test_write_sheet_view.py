###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...worksheet import Worksheet


class TestWriteSheetView(unittest.TestCase):
    """
    Test the Worksheet _write_sheet_view() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_write_sheet_view_tab_not_selected(self):
        """Test the _write_sheet_view() method. Tab not selected"""

        self.worksheet._write_sheet_view()

        exp = """<sheetView workbookViewId="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_view_tab_selected(self):
        """Test the _write_sheet_view() method. Tab selected"""

        self.worksheet.select()
        self.worksheet._write_sheet_view()

        exp = """<sheetView tabSelected="1" workbookViewId="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_view_hide_gridlines(self):
        """Test the _write_sheet_view() method. Tab selected + hide_gridlines()"""

        self.worksheet.select()
        self.worksheet.hide_gridlines()
        self.worksheet._write_sheet_view()

        exp = """<sheetView tabSelected="1" workbookViewId="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_view_hide_gridlines_0(self):
        """Test the _write_sheet_view() method. Tab selected + hide_gridlines(0)"""

        self.worksheet.select()
        self.worksheet.hide_gridlines(0)
        self.worksheet._write_sheet_view()

        exp = """<sheetView tabSelected="1" workbookViewId="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_view_hide_gridlines_1(self):
        """Test the _write_sheet_view() method. Tab selected + hide_gridlines(1)"""

        self.worksheet.select()
        self.worksheet.hide_gridlines(1)
        self.worksheet._write_sheet_view()

        exp = """<sheetView tabSelected="1" workbookViewId="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_view_hide_gridlines_2(self):
        """Test the _write_sheet_view() method. Tab selected + hide_gridlines(2)"""

        self.worksheet.select()
        self.worksheet.hide_gridlines(2)
        self.worksheet._write_sheet_view()

        exp = """<sheetView showGridLines="0" tabSelected="1" workbookViewId="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
