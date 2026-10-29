###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import unittest

from xlsxwriter.utility import cell_autofit_width, xl_pixel_width


class TestUtility(unittest.TestCase):
    """
    Test xl_pixel_width() utility function.

    """

    def assert_pixel_width(self, string, exp):
        """Check xl_pixel_width() and cell_autofit_width() for a string."""
        self.assertEqual(exp, xl_pixel_width(string))
        self.assertEqual(cell_autofit_width(string), exp + 7)

    def test_xl_pixel_width_punctuation(self):
        """Test xl_pixel_width() with space and punctuation"""

        self.assert_pixel_width(" ", 3)
        self.assert_pixel_width("!", 5)
        self.assert_pixel_width('"', 6)
        self.assert_pixel_width("#", 7)
        self.assert_pixel_width("$", 7)
        self.assert_pixel_width("%", 11)
        self.assert_pixel_width("&", 10)
        self.assert_pixel_width("'", 3)
        self.assert_pixel_width("(", 5)
        self.assert_pixel_width(")", 5)
        self.assert_pixel_width("*", 7)
        self.assert_pixel_width("+", 7)
        self.assert_pixel_width(",", 4)
        self.assert_pixel_width("-", 5)
        self.assert_pixel_width(".", 4)
        self.assert_pixel_width("/", 6)
        self.assert_pixel_width(":", 4)
        self.assert_pixel_width(";", 4)
        self.assert_pixel_width("<", 7)
        self.assert_pixel_width("=", 7)
        self.assert_pixel_width(">", 7)
        self.assert_pixel_width("?", 7)
        self.assert_pixel_width("@", 13)
        self.assert_pixel_width("[", 5)
        self.assert_pixel_width("\\", 6)
        self.assert_pixel_width("]", 5)
        self.assert_pixel_width("^", 7)
        self.assert_pixel_width("_", 7)
        self.assert_pixel_width("`", 4)
        self.assert_pixel_width("{", 5)
        self.assert_pixel_width("|", 7)
        self.assert_pixel_width("}", 5)
        self.assert_pixel_width("~", 7)

    def test_xl_pixel_width_digits(self):
        """Test xl_pixel_width() with digits"""

        self.assert_pixel_width("0", 7)
        self.assert_pixel_width("1", 7)
        self.assert_pixel_width("2", 7)
        self.assert_pixel_width("3", 7)
        self.assert_pixel_width("4", 7)
        self.assert_pixel_width("5", 7)
        self.assert_pixel_width("6", 7)
        self.assert_pixel_width("7", 7)
        self.assert_pixel_width("8", 7)
        self.assert_pixel_width("9", 7)

    def test_xl_pixel_width_uppercase(self):
        """Test xl_pixel_width() with uppercase letters"""

        self.assert_pixel_width("A", 9)
        self.assert_pixel_width("B", 8)
        self.assert_pixel_width("C", 8)
        self.assert_pixel_width("D", 9)
        self.assert_pixel_width("E", 7)
        self.assert_pixel_width("F", 7)
        self.assert_pixel_width("G", 9)
        self.assert_pixel_width("H", 9)
        self.assert_pixel_width("I", 4)
        self.assert_pixel_width("J", 5)
        self.assert_pixel_width("K", 8)
        self.assert_pixel_width("L", 6)
        self.assert_pixel_width("M", 12)
        self.assert_pixel_width("N", 10)
        self.assert_pixel_width("O", 10)
        self.assert_pixel_width("P", 8)
        self.assert_pixel_width("Q", 10)
        self.assert_pixel_width("R", 8)
        self.assert_pixel_width("S", 7)
        self.assert_pixel_width("T", 7)
        self.assert_pixel_width("U", 9)
        self.assert_pixel_width("V", 9)
        self.assert_pixel_width("W", 13)
        self.assert_pixel_width("X", 8)
        self.assert_pixel_width("Y", 7)
        self.assert_pixel_width("Z", 7)

    def test_xl_pixel_width_lowercase(self):
        """Test xl_pixel_width() with lowercase letters"""

        self.assert_pixel_width("a", 7)
        self.assert_pixel_width("b", 8)
        self.assert_pixel_width("c", 6)
        self.assert_pixel_width("d", 8)
        self.assert_pixel_width("e", 8)
        self.assert_pixel_width("f", 5)
        self.assert_pixel_width("g", 7)
        self.assert_pixel_width("h", 8)
        self.assert_pixel_width("i", 4)
        self.assert_pixel_width("j", 4)
        self.assert_pixel_width("k", 7)
        self.assert_pixel_width("l", 4)
        self.assert_pixel_width("m", 12)
        self.assert_pixel_width("n", 8)
        self.assert_pixel_width("o", 8)
        self.assert_pixel_width("p", 8)
        self.assert_pixel_width("q", 8)
        self.assert_pixel_width("r", 5)
        self.assert_pixel_width("s", 6)
        self.assert_pixel_width("t", 5)
        self.assert_pixel_width("u", 8)
        self.assert_pixel_width("v", 7)
        self.assert_pixel_width("w", 11)
        self.assert_pixel_width("x", 7)
        self.assert_pixel_width("y", 7)
        self.assert_pixel_width("z", 6)

    def test_xl_pixel_width_unicode_and_strings(self):
        """Test xl_pixel_width() with unicode and multi-char strings"""

        self.assert_pixel_width("é", 8)
        self.assert_pixel_width("éé", 16)
        self.assert_pixel_width("ABC", 25)
        self.assert_pixel_width("Hello", 33)
        self.assert_pixel_width("12345", 35)
