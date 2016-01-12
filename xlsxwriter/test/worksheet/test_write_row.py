###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...worksheet import Worksheet
from ...format import Format


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

        self.assertEqual(got, exp)

    def test_write_row_2(self):
        """Test the _write_row() method"""

        self.worksheet._write_row(2, '2:2')

        exp = """<row r="3" spans="2:2">"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_row_3(self):
        """Test the _write_row() method"""

        self.worksheet._write_row(1, None, [30, None, 0, 0, 0])

        exp = """<row r="2" ht="30" customHeight="1">"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_row_4(self):
        """Test the _write_row() method"""

        self.worksheet._write_row(3, None, [15, None, 1, 0, 0])

        exp = """<row r="4" hidden="1">"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_row_5(self):
        """Test the _write_row() method"""

        cell_format = Format({'xf_index': 1})

        self.worksheet._write_row(6, None, [15, cell_format, 0, 0, 0])

        exp = """<row r="7" s="1" customFormat="1">"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_row_6(self):
        """Test the _write_row() method"""

        self.worksheet._write_row(9, None, [3, None, 0, 0, 0])

        exp = """<row r="10" ht="3" customHeight="1">"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_row_7(self):
        """Test the _write_row() method"""

        self.worksheet._write_row(12, None, [24, None, 1, 0, 0])

        exp = """<row r="13" ht="24" hidden="1" customHeight="1">"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_row_8(self):
        """Test the _write_row() method"""

        self.worksheet._write_row(12, None, [24, None, 1, 0, 0], 1)

        exp = """<row r="13" ht="24" hidden="1" customHeight="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
