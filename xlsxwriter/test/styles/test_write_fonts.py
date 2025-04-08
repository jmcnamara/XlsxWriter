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

from ...format import Format
from ...styles import Styles


class TestWriteFonts(unittest.TestCase):
    """
    Test the Styles _write_fonts() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.styles = Styles()
        self.styles._set_filehandle(self.fh)

    def test_write_fonts(self):
        """Test the _write_fonts() method"""

        xf_format = Format()
        xf_format.has_font = True

        self.styles._set_style_properties([[xf_format], None, 1, 0, 0, 0, [], [], 0])

        self.styles._write_fonts()

        exp = """<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>"""
        got = self.fh.getvalue()

        self.assertEqual(exp, got)
