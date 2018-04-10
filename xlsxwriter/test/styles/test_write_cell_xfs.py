###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...styles import Styles
from ...format import Format


class TestWriteCellXfs(unittest.TestCase):
    """
    Test the Styles _write_cell_xfs() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.styles = Styles()
        self.styles._set_filehandle(self.fh)

    def test_write_cell_xfs(self):
        """Test the _write_cell_xfs() method"""

        xf_format = Format()
        xf_format.has_font = 1

        self.styles._set_style_properties([[xf_format], None, 1, 0, 0, 0, [], []])

        self.styles._write_cell_xfs()

        exp = """<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
