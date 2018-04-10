###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...drawing import Drawing


class TestWriteXdrext(unittest.TestCase):
    """
    Test the Drawing _write_ext() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.drawing = Drawing()
        self.drawing._set_filehandle(self.fh)

    def test_write_ext(self):
        """Test the _write_ext() method"""

        self.drawing._write_ext(9308969, 6078325)

        exp = """<xdr:ext cx="9308969" cy="6078325"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
