###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ...worksheet import Worksheet


class TestWritePageMargins(unittest.TestCase):
    """
    Test the Worksheet _write_page_margins() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_write_page_margins(self):
        """Test the _write_page_margins() method"""

        self.worksheet._write_page_margins()

        exp = """<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_page_margins_deafult(self):
        """Test the _write_page_margins() method with default margins"""

        self.worksheet.set_margins()
        self.worksheet._write_page_margins()

        exp = """<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_page_margins_left(self):
        """Test the _write_page_margins() method with left margin"""

        self.worksheet.set_margins(left=0.5)
        self.worksheet._write_page_margins()

        exp = """<pageMargins left="0.5" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_page_margins_right(self):
        """Test the _write_page_margins() method with right margin"""

        self.worksheet.set_margins(right=0.5)
        self.worksheet._write_page_margins()

        exp = """<pageMargins left="0.7" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_page_margins_top(self):
        """Test the _write_page_margins() method with top margin"""

        self.worksheet.set_margins(top=0.5)
        self.worksheet._write_page_margins()

        exp = """<pageMargins left="0.7" right="0.7" top="0.5" bottom="0.75" header="0.3" footer="0.3"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_page_margins_bottom(self):
        """Test the _write_page_margins() method with bottom margin"""

        self.worksheet.set_margins(bottom=0.5)
        self.worksheet._write_page_margins()

        exp = """<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.5" header="0.3" footer="0.3"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_page_margins_header(self):
        """Test the _write_page_margins() method with header margin"""

        self.worksheet.set_header(margin=0.5)
        self.worksheet._write_page_margins()

        exp = """<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.5" footer="0.3"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_page_margins_footer(self):
        """Test the _write_page_margins() method with footer margin"""

        self.worksheet.set_footer(margin=0.5)
        self.worksheet._write_page_margins()

        exp = """<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.5"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
