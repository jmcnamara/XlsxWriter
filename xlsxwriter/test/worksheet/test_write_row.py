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

from xlsxwriter.format import Format
from xlsxwriter.worksheet import RowInfo, Worksheet


class TestWriteRow(unittest.TestCase):
    """
    Test the Worksheet _write_row() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_write_row_1(self):
        """Test the _write_row() method"""

        self.worksheet._write_row(0, None)

        exp = """<row r="1">"""
        got = self.fh.getvalue()

        self.assertEqual(exp, got)

    def test_write_row_2(self):
        """Test the _write_row() method"""

        self.worksheet._write_row(2, "2:2")

        exp = """<row r="3" spans="2:2">"""
        got = self.fh.getvalue()

        self.assertEqual(exp, got)

    def test_write_row_3(self):
        """Test the _write_row() method"""

        row_info = RowInfo(height=30)
        self.worksheet._write_row(1, None, row_info)

        exp = """<row r="2" ht="30" customHeight="1">"""
        got = self.fh.getvalue()

        self.assertEqual(exp, got)

    def test_write_row_4(self):
        """Test the _write_row() method"""

        row_info = RowInfo(height=15, hidden=True)
        self.worksheet._write_row(3, None, row_info)

        exp = """<row r="4" hidden="1">"""
        got = self.fh.getvalue()

        self.assertEqual(exp, got)

    def test_write_row_5(self):
        """Test the _write_row() method"""

        row_format = Format({"xf_index": 1})
        row_info = RowInfo(height=15, row_format=row_format)

        self.worksheet._write_row(6, None, row_info)

        exp = """<row r="7" s="1" customFormat="1">"""
        got = self.fh.getvalue()

        self.assertEqual(exp, got)

    def test_write_row_6(self):
        """Test the _write_row() method"""

        row_info = RowInfo(height=3)
        self.worksheet._write_row(9, None, row_info)

        exp = """<row r="10" ht="3" customHeight="1">"""
        got = self.fh.getvalue()

        self.assertEqual(exp, got)

    def test_write_row_7(self):
        """Test the _write_row() method"""

        row_info = RowInfo(height=24, hidden=True)
        self.worksheet._write_row(12, None, row_info)

        exp = """<row r="13" ht="24" hidden="1" customHeight="1">"""
        got = self.fh.getvalue()

        self.assertEqual(exp, got)

    def test_write_row_8(self):
        """Test the _write_row() method"""

        row_info = RowInfo(height=24, hidden=True)
        self.worksheet._write_row(12, None, row_info, 1)

        exp = """<row r="13" ht="24" hidden="1" customHeight="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(exp, got)
