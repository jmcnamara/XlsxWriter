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

from xlsxwriter.worksheet import Worksheet


class TestWriteTabColor(unittest.TestCase):
    """
    Test the Worksheet _write_tab_color() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_write_tab_color(self):
        """Test the _write_tab_color() method"""

        self.worksheet.set_tab_color("red")
        self.worksheet._write_tab_color()

        exp = """<tabColor rgb="FFFF0000"/>"""
        got = self.fh.getvalue()

        self.assertEqual(exp, got)
