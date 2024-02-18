###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ...vml import Vml


class TestWriteVfill(unittest.TestCase):
    """
    Test the Vml _write_fill() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.vml = Vml()
        self.vml._set_filehandle(self.fh)

    def test_write_comment_fill(self):
        """Test the _write_comment_fill() method"""

        self.vml._write_comment_fill()

        exp = """<v:fill color2="#ffffe1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_button_fill(self):
        """Test the _write_button_fill() method"""

        self.vml._write_button_fill()

        exp = """<v:fill color2="buttonFace [67]" o:detectmouseclick="t"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
