###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...drawing import Drawing


class TestWriteXdrcolOff(unittest.TestCase):
    """
    Test the Drawing _write_col_off() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.drawing = Drawing()
        self.drawing._set_filehandle(self.fh)

    def test_write_col_off(self):
        """Test the _write_col_off() method"""

        self.drawing._write_col_off(457200)

        exp = """<xdr:colOff>457200</xdr:colOff>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
