###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...styles import Styles


class TestWriteCellStyleXfs(unittest.TestCase):
    """
    Test the Styles _write_cell_style_xfs() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.styles = Styles()
        self.styles._set_filehandle(self.fh)

    def test_write_cell_style_xfs(self):
        """Test the _write_cell_style_xfs() method"""

        self.styles._write_cell_style_xfs()

        exp = """<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
