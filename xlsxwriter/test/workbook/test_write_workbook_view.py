###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...workbook import Workbook


class TestWriteWorkbookView(unittest.TestCase):
    """
    Test the Workbook _write_workbook_view() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.workbook = Workbook()
        self.workbook._set_filehandle(self.fh)

    def test_write_workbook_view1(self):
        """Test the _write_workbook_view() method"""

        self.workbook._write_workbook_view()

        exp = """<workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_workbook_view2(self):
        """Test the _write_workbook_view() method"""

        self.workbook.worksheet_meta.activesheet = 1
        self.workbook._write_workbook_view()

        exp = """<workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660" activeTab="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_workbook_view3(self):
        """Test the _write_workbook_view() method"""

        self.workbook.worksheet_meta.firstsheet = 1
        self.workbook.worksheet_meta.activesheet = 1
        self.workbook._write_workbook_view()

        exp = """<workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660" firstSheet="2" activeTab="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_workbook_view4(self):
        """Test the _write_workbook_view() method"""

        self.workbook.set_size(0, 0)
        self.workbook._write_workbook_view()

        exp = """<workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_workbook_view5(self):
        """Test the _write_workbook_view() method"""

        self.workbook.set_size(None, None)
        self.workbook._write_workbook_view()

        exp = """<workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_workbook_view5(self):
        """Test the _write_workbook_view() method"""

        self.workbook.set_size(1073, 644)
        self.workbook._write_workbook_view()

        exp = """<workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_workbook_view6(self):
        """Test the _write_workbook_view() method"""

        self.workbook.set_size(123, 70)
        self.workbook._write_workbook_view()

        exp = """<workbookView xWindow="240" yWindow="15" windowWidth="1845" windowHeight="1050"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_workbook_view7(self):
        """Test the _write_workbook_view() method"""

        self.workbook.set_size(719, 490)
        self.workbook._write_workbook_view()

        exp = """<workbookView xWindow="240" yWindow="15" windowWidth="10785" windowHeight="7350"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def tearDown(self):
        self.workbook.fileclosed = 1
