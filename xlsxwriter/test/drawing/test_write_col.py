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


class TestWriteXdrcol(unittest.TestCase):
    """
    Test the Drawing _write_col() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.drawing = Drawing()
        self.drawing._set_filehandle(self.fh)

    def test_write_col(self):
        """Test the _write_col() method"""

        self.drawing._write_col(4)

        exp = """<xdr:col>4</xdr:col>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
