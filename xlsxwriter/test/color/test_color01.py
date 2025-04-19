###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#


import unittest

from xlsxwriter.color import Color, ColorTypes


class TestColor(unittest.TestCase):
    """
    Test cases for the Color class.
    """

    def test_color_rgb_from_string(self):
        """Test creating a Color instance from a hex string."""
        color = Color("#FF5733")
        self.assertEqual(color._rgb_value, 0xFF5733)
        self.assertEqual(color._type, ColorTypes.RGB)
        self.assertFalse(color._is_automatic)

    def test_color_rgb_from_int(self):
        """Test creating a Color instance from an integer RGB value."""
        color = Color(0x00FF00)
        self.assertEqual(color._rgb_value, 0x00FF00)
        self.assertEqual(color._type, ColorTypes.RGB)
        self.assertFalse(color._is_automatic)

    def test_color_theme(self):
        """Test creating a Color instance from a theme color tuple."""
        color = Color((2, 3))
        self.assertEqual(color._theme_color, (2, 3))
        self.assertEqual(color._type, ColorTypes.THEME)
        self.assertFalse(color._is_automatic)

    def test_color_invalid_string(self):
        """Test creating a Color instance with an invalid string."""
        with self.assertRaises(ValueError):
            Color("invalid")

    def test_color_invalid_int(self):
        """Test creating a Color instance with an out-of-range integer."""
        with self.assertRaises(ValueError):
            Color(0xFFFFFF + 1)

    def test_color_invalid_theme(self):
        """Test creating a Color instance with an invalid theme tuple."""
        with self.assertRaises(ValueError):
            Color((10, 2))  # Invalid theme color
        with self.assertRaises(ValueError):
            Color((2, 6))  # Invalid theme shade

    def test_is_automatic_property(self):
        """Test setting and getting the is_automatic property."""
        color = Color("#000000")
        color._is_automatic = True
        self.assertTrue(color._is_automatic)
