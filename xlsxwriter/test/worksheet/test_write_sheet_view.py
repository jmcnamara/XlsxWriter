###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013, John McNamara, jmcnamara@cpan.org
#

import unittest
from StringIO import StringIO
from ...worksheet import Worksheet


class TestWriteSheetView(unittest.TestCase):
    """
    Test the Worksheet _write_sheet_view() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_write_sheet_view(self):
        """Test the _write_sheet_view() method"""

        self.worksheet.selected = 1
        self.worksheet._write_sheet_view()

        exp = """<sheetView tabSelected="1" workbookViewId="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)


if __name__ == '__main__':
    unittest.main()
