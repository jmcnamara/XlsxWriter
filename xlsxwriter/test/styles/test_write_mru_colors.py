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


class TestWriteMruColors(unittest.TestCase):
    """
    Test the Styles _write_mru_colors() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.styles = Styles()
        self.styles._set_filehandle(self.fh)

    def test_write_mru_colors(self):
        """Test the _write_mru_colors() method"""

        self.styles._write_mru_colors(["FF26DA55", "FF792DC8", "FF646462"])

        exp = """<mruColors><color rgb="FF646462"/><color rgb="FF792DC8"/><color rgb="FF26DA55"/></mruColors>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
