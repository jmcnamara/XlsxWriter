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


class TestPixelsToRowCol(unittest.TestCase):
    """
    Test the Worksheet _pixels_to_xxx() methods.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    # Function for testing.
    def width_to_pixels(self, width):
        max_digit_width = 7
        padding = 5

        if width < 1:
            pixels = int(width * (max_digit_width + padding) + 0.5)
        else:
            pixels = int(width * max_digit_width + 0.5) + padding

        return pixels

    # Function for testing.
    def height_to_pixels(self, height):
        return int(4.0 / 3.0 * height)

    def test_pixels_to_width(self):
        """Test the _pixels_to_width() function"""

        for pixels in range(1791):
            exp = pixels
            got = self.width_to_pixels(self.worksheet._pixels_to_width(pixels))

            self.assertEqual(got, exp)

    def test_pixels_to_height(self):
        """Test the _pixels_to_height() function"""

        for pixels in range(546):
            exp = pixels
            got = self.height_to_pixels(self.worksheet._pixels_to_height(pixels))

            self.assertEqual(got, exp)
