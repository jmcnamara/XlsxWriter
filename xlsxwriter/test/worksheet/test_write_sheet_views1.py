###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
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

    def test_write_sheet_views(self):
        """Test the _write_sheet_views() method"""

        self.worksheet.select()
        self.worksheet._write_sheet_views()

        exp = """<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_views_zoom_100(self):
        """Test the _write_sheet_views() method"""

        self.worksheet.select()
        self.worksheet.set_zoom(100)  # Default. Should be ignored.
        self.worksheet._write_sheet_views()

        exp = """<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_views_zoom_200(self):
        """Test the _write_sheet_views() method"""

        self.worksheet.select()
        self.worksheet.set_zoom(200)
        self.worksheet._write_sheet_views()

        exp = """<sheetViews><sheetView tabSelected="1" zoomScale="200" zoomScaleNormal="200" workbookViewId="0"/></sheetViews>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_views_right_to_left(self):
        """Test the _write_sheet_views() method"""

        self.worksheet.select()
        self.worksheet.right_to_left()
        self.worksheet._write_sheet_views()

        exp = """<sheetViews><sheetView rightToLeft="1" tabSelected="1" workbookViewId="0"/></sheetViews>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_views_hide_zero(self):
        """Test the _write_sheet_views() method"""

        self.worksheet.select()
        self.worksheet.hide_zero()
        self.worksheet._write_sheet_views()

        exp = """<sheetViews><sheetView showZeros="0" tabSelected="1" workbookViewId="0"/></sheetViews>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_views_page_view(self):
        """Test the _write_sheet_views() method"""

        self.worksheet.select()
        self.worksheet.set_page_view()
        self.worksheet._write_sheet_views()

        exp = """<sheetViews><sheetView tabSelected="1" view="pageLayout" workbookViewId="0"/></sheetViews>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
