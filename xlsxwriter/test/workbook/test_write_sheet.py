###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...workbook import Workbook


class TestWriteSheet(unittest.TestCase):
    """
    Test the Workbook _write_sheet() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.workbook = Workbook()
        self.workbook._set_filehandle(self.fh)

    def test_write_sheet1(self):
        """Test the _write_sheet() method"""

        self.workbook._write_sheet('Sheet1', 1, 0)

        exp = """<sheet name="Sheet1" sheetId="1" r:id="rId1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet2(self):
        """Test the _write_sheet() method"""

        self.workbook._write_sheet('Sheet1', 1, 1)

        exp = """<sheet name="Sheet1" sheetId="1" state="hidden" r:id="rId1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_sheet3(self):
        """Test the _write_sheet() method"""

        self.workbook._write_sheet('Bits & Bobs', 1, 0)

        exp = """<sheet name="Bits &amp; Bobs" sheetId="1" r:id="rId1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def tearDown(self):
        self.workbook.fileclosed = 1
