###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2022, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ...drawing import Drawing


class TestWriteXdrrow(unittest.TestCase):
    """
    Test the Drawing _write_row() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.drawing = Drawing()
        self.drawing._set_filehandle(self.fh)

    def test_write_row(self):
        """Test the _write_row() method"""

        self.drawing._write_row(8)

        exp = """<xdr:row>8</xdr:row>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
