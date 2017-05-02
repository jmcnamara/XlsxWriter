###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...workbook import Workbook


class TestDeleteSheet(unittest.TestCase):
    """
    Test the Workbook _delete_sheet() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.workbook = Workbook()
        self.workbook._set_filehandle(self.fh)

    def test_delete_sheet(self):
        """Test the _delete_sheets() method"""

        self.workbook.add_worksheet('Sheet1')
        self.workbook.add_worksheet('Sheet2')
        self.workbook._write_sheets()

        exp = """<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rId2"/></sheets>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
        self.fh.close()
        self.fh = StringIO()
        self.workbook._set_filehandle(self.fh)

        self.workbook._remove_sheet('Sheet1')
        self.workbook._write_sheets()

        exp = """<sheets><sheet name="Sheet2" sheetId="1" r:id="rId1"/></sheets>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def tearDown(self):
        self.workbook.fileclosed = 1
