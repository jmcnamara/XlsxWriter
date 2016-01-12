###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...workbook import Workbook


class TestWriteBookViews(unittest.TestCase):
    """
    Test the Workbook _write_book_views() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.workbook = Workbook()
        self.workbook._set_filehandle(self.fh)

    def test_write_book_views(self):
        """Test the _write_book_views() method"""

        self.workbook._write_book_views()

        exp = """<bookViews><workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/></bookViews>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def tearDown(self):
        self.workbook.fileclosed = 1
