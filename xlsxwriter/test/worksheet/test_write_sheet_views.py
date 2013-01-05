###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013, John McNamara, jmcnamara@cpan.org
#

import unittest
from StringIO import StringIO
from ...worksheet import Worksheet


class TestWriteSheetViews(unittest.TestCase):
    """
    Test the Worksheet _write_sheet_views() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet(self.fh)

    def test_write_sheet_views(self):
        """Test the _write_sheet_views() method"""

        self.worksheet._write_sheet_views()

        exp = """<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)


if __name__ == '__main__':
    unittest.main()
