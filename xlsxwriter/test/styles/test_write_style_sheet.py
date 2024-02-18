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


class TestWriteStyleSheet(unittest.TestCase):
    """
    Test the Styles _write_style_sheet() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.styles = Styles()
        self.styles._set_filehandle(self.fh)

    def test_write_style_sheet(self):
        """Test the _write_style_sheet() method"""

        self.styles._write_style_sheet()

        exp = """<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
