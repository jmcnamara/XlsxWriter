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

from xlsxwriter.styles import Styles


class TestWriteStyleXf(unittest.TestCase):
    """
    Test the Styles _write_style_xf() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.styles = Styles()
        self.styles._set_filehandle(self.fh)

    def test_write_style_xf(self):
        """Test the _write_style_xf() method"""

        self.styles._write_style_xf()

        exp = """<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(exp, got)
