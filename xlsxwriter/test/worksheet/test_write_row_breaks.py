###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...worksheet import Worksheet


class TestWriteRowBreaks(unittest.TestCase):
    """
    Test the Worksheet _write_row_breaks() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_write_row_breaks_1(self):
        """Test the _write_row_breaks() method"""

        self.worksheet.hbreaks = [1]

        self.worksheet._write_row_breaks()

        exp = """<rowBreaks count="1" manualBreakCount="1"><brk id="1" max="16383" man="1"/></rowBreaks>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_row_breaks_2(self):
        """Test the _write_row_breaks() method"""

        self.worksheet.hbreaks = [15, 7, 3, 0]

        self.worksheet._write_row_breaks()

        exp = """<rowBreaks count="3" manualBreakCount="3"><brk id="3" max="16383" man="1"/><brk id="7" max="16383" man="1"/><brk id="15" max="16383" man="1"/></rowBreaks>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
