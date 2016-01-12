###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...worksheet import Worksheet


class TestWriteSheetPr(unittest.TestCase):
    """
    Test the Worksheet _write_sheet_pr() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_write_sheet_pr_fit_to_page(self):
        """Test the _write_sheet_pr() method"""

        self.worksheet.fit_to_pages(1, 1)
        self.worksheet._write_sheet_pr()

        exp = """<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_pr_tab_color(self):
        """Test the _write_sheet_pr() method"""

        self.worksheet.set_tab_color('red')
        self.worksheet._write_sheet_pr()

        exp = """<sheetPr><tabColor rgb="FFFF0000"/></sheetPr>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet_pr_both(self):
        """Test the _write_sheet_pr() method"""

        self.worksheet.set_tab_color('red')
        self.worksheet.fit_to_pages(1, 1)
        self.worksheet._write_sheet_pr()

        exp = """<sheetPr><tabColor rgb="FFFF0000"/><pageSetUpPr fitToPage="1"/></sheetPr>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
