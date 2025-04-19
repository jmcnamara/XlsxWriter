###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO

from xlsxwriter.color import Color
from xlsxwriter.styles import Styles


class TestWriteColors(unittest.TestCase):
    """
    Test the Styles _write_colors() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.styles = Styles()
        self.styles._set_filehandle(self.fh)

    def test_write_colors1(self):
        """Test the _write_colors() method"""

        self.styles.custom_colors = [Color("#26DA55")]
        self.styles._write_colors()

        exp = """<colors><mruColors><color rgb="FF26DA55"/></mruColors></colors>"""
        got = self.fh.getvalue()

        self.assertEqual(exp, got)

    def test_write_colors2(self):
        """Test the _write_colors() method"""

        self.styles.custom_colors = [
            Color("#26DA55"),
            Color("#792DC8"),
            Color("#646462"),
        ]
        self.styles._write_colors()

        exp = """<colors><mruColors><color rgb="FF646462"/><color rgb="FF792DC8"/><color rgb="FF26DA55"/></mruColors></colors>"""
        got = self.fh.getvalue()

        self.assertEqual(exp, got)

    def test_write_colors3(self):
        """Test the _write_colors() method"""

        self.styles.custom_colors = [
            Color("#792DC8"),
            Color("#646462"),
            Color("#5EA29C"),
            Color("#583AC6"),
            Color("#E31DAF"),
            Color("#A1A759"),
            Color("#600FF1"),
            Color("#0CF49C"),
            Color("#E3FA06"),
            Color("#913AC6"),
            Color("#B97847"),
            Color("#D97827"),
        ]

        self.styles._write_colors()

        exp = """<colors><mruColors><color rgb="FFD97827"/><color rgb="FFB97847"/><color rgb="FF913AC6"/><color rgb="FFE3FA06"/><color rgb="FF0CF49C"/><color rgb="FF600FF1"/><color rgb="FFA1A759"/><color rgb="FFE31DAF"/><color rgb="FF583AC6"/><color rgb="FF5EA29C"/></mruColors></colors>"""
        got = self.fh.getvalue()

        self.assertEqual(exp, got)
