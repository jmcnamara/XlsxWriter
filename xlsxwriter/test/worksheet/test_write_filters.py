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


class TestWriteFilters(unittest.TestCase):
    """
    Test the Worksheet _write_filters() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_write_filters_1(self):
        """Test the _write_filters() method"""

        self.worksheet._write_filters(["East"])

        exp = """<filters><filter val="East"/></filters>"""
        got = self.fh.getvalue()

        self.assertEqual(exp, got)

    def test_write_filters_2(self):
        """Test the _write_filters() method"""

        self.worksheet._write_filters(["East", "South"])

        exp = """<filters><filter val="East"/><filter val="South"/></filters>"""
        got = self.fh.getvalue()

        self.assertEqual(exp, got)

    def test_write_filters_3(self):
        """Test the _write_filters() method"""

        self.worksheet._write_filters(["blanks"])

        exp = """<filters blank="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(exp, got)
