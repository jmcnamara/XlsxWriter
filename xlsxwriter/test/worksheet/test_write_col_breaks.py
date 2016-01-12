###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...worksheet import Worksheet


class TestWriteColBreaks(unittest.TestCase):
    """
    Test the Worksheet _write_col_breaks() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_write_col_breaks_1(self):
        """Test the _write_col_breaks() method"""

        self.worksheet.vbreaks = [1]

        self.worksheet._write_col_breaks()

        exp = """<colBreaks count="1" manualBreakCount="1"><brk id="1" max="1048575" man="1"/></colBreaks>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_col_breaks_2(self):
        """Test the _write_col_breaks() method"""

        self.worksheet.vbreaks = [8, 3, 1, 0]

        self.worksheet._write_col_breaks()

        exp = """<colBreaks count="3" manualBreakCount="3"><brk id="1" max="1048575" man="1"/><brk id="3" max="1048575" man="1"/><brk id="8" max="1048575" man="1"/></colBreaks>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
