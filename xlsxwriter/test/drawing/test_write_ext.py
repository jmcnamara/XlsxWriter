###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ...drawing import Drawing


class TestWriteXdrext(unittest.TestCase):
    """
    Test the Drawing _write_ext() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.drawing = Drawing()
        self.drawing._set_filehandle(self.fh)

    def test_write_xdr_ext(self):
        """Test the _write_ext() method"""

        self.drawing._write_xdr_ext(9308969, 6078325)

        exp = """<xdr:ext cx="9308969" cy="6078325"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
