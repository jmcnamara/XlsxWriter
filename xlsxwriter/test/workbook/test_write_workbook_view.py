###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013, John McNamara, jmcnamara@cpan.org
#

import unittest
from ..compatibility import StringIO
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

        exp = """<workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660" firstSheet="1" activeTab="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def tearDown(self):
        self.workbook.fileclosed = 1


if __name__ == '__main__':
    unittest.main()
