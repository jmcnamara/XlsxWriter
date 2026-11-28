###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO

from xlsxwriter.drawing import Drawing
from xlsxwriter.shape import Shape


class TestWriteFontRun(unittest.TestCase):
    """
    Test the Drawing _write_font_run() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.drawing = Drawing()
        self.drawing._set_filehandle(self.fh)

    def test_write_font_run_fractional_size(self):
        """Test the _write_font_run() method with a fractional font size"""

        font = Shape._get_font_properties({"size": 1.13})
        style_attrs = Shape._get_font_style_attributes(font)
        style_attrs.insert(0, ("lang", font["lang"]))

        self.drawing._write_font_run(font, style_attrs, [], "a:rPr")

        exp = """<a:rPr lang="en-US" sz="113"/>"""
        got = self.fh.getvalue()

        self.assertEqual(exp, got)
