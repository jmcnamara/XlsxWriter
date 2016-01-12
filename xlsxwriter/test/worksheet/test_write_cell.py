###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...compat_collections import namedtuple
from ...worksheet import Worksheet


class TestWriteCell(unittest.TestCase):
    """
    Test the Worksheet _write_cell() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_write_cell_number(self):
        """Test the _write_cell() method for numbers."""

        cell_tuple = namedtuple('Number', 'number, format')
        cell = cell_tuple(1, None)

        self.worksheet._write_cell(0, 0, cell)

        exp = """<c r="A1"><v>1</v></c>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_cell_string(self):
        """Test the _write_cell() method for strings."""

        cell_tuple = namedtuple('String', 'string, format')
        cell = cell_tuple(0, None)

        self.worksheet._write_cell(3, 1, cell)

        exp = """<c r="B4" t="s"><v>0</v></c>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_cell_formula01(self):
        """Test the _write_cell() method for formulas."""

        cell_tuple = namedtuple('Formula', 'formula, format, value')
        cell = cell_tuple('A3+A5', None, 0)

        self.worksheet._write_cell(1, 2, cell)

        exp = """<c r="C2"><f>A3+A5</f><v>0</v></c>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_cell_formula02(self):
        """Test the _write_cell() method for formulas."""

        cell_tuple = namedtuple('Formula', 'formula, format, value')
        cell = cell_tuple('A3+A5', None, 7)

        self.worksheet._write_cell(1, 2, cell)

        exp = """<c r="C2"><f>A3+A5</f><v>7</v></c>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
