###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...drawing import Drawing


class TestWriteXdrpos(unittest.TestCase):
    """
    Test the Drawing _write_pos() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.drawing = Drawing()
        self.drawing._set_filehandle(self.fh)

    def test_write_pos(self):
        """Test the _write_pos() method"""

        self.drawing._write_pos(0, 0)

        exp = """<xdr:pos x="0" y="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
