###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ...styles import Styles


class TestWriteCellStyles(unittest.TestCase):
    """
    Test the Styles _write_cell_styles() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.styles = Styles()
        self.styles._set_filehandle(self.fh)

    def test_write_cell_styles(self):
        """Test the _write_cell_styles() method"""

        self.styles._write_cell_styles()

        exp = """<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
