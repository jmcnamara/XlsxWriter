###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

import unittest

from xlsxwriter.utility import cell_autofit_width, xl_pixel_width


class TestUtility(unittest.TestCase):
    """
    Test xl_pixel_width() utility function.

    """

    def test_xl_pixel_width(self):
        """Test xl_pixel_width()"""

        tests = [
            (" ", 3),
            ("!", 5),
            ('"', 6),
            ("#", 7),
            ("$", 7),
            ("%", 11),
            ("&", 10),
            ("'", 3),
            ("(", 5),
            (")", 5),
            ("*", 7),
            ("+", 7),
            (",", 4),
            ("-", 5),
            (".", 4),
            ("/", 6),
            ("0", 7),
            ("1", 7),
            ("2", 7),
            ("3", 7),
            ("4", 7),
            ("5", 7),
            ("6", 7),
            ("7", 7),
            ("8", 7),
            ("9", 7),
            (":", 4),
            (";", 4),
            ("<", 7),
            ("=", 7),
            (">", 7),
            ("?", 7),
            ("@", 13),
            ("A", 9),
            ("B", 8),
            ("C", 8),
            ("D", 9),
            ("E", 7),
            ("F", 7),
            ("G", 9),
            ("H", 9),
            ("I", 4),
            ("J", 5),
            ("K", 8),
            ("L", 6),
            ("M", 12),
            ("N", 10),
            ("O", 10),
            ("P", 8),
            ("Q", 10),
            ("R", 8),
            ("S", 7),
            ("T", 7),
            ("U", 9),
            ("V", 9),
            ("W", 13),
            ("X", 8),
            ("Y", 7),
            ("Z", 7),
            ("[", 5),
            ("\\", 6),
            ("]", 5),
            ("^", 7),
            ("_", 7),
            ("`", 4),
            ("a", 7),
            ("b", 8),
            ("c", 6),
            ("d", 8),
            ("e", 8),
            ("f", 5),
            ("g", 7),
            ("h", 8),
            ("i", 4),
            ("j", 4),
            ("k", 7),
            ("l", 4),
            ("m", 12),
            ("n", 8),
            ("o", 8),
            ("p", 8),
            ("q", 8),
            ("r", 5),
            ("s", 6),
            ("t", 5),
            ("u", 8),
            ("v", 7),
            ("w", 11),
            ("x", 7),
            ("y", 7),
            ("z", 6),
            ("{", 5),
            ("|", 7),
            ("}", 5),
            ("~", 7),
            ("é", 8),
            ("éé", 16),
            ("ABC", 25),
            ("Hello", 33),
            ("12345", 35),
        ]

        for string, exp in tests:
            got = xl_pixel_width(string)
            self.assertEqual(exp, got)

            got = cell_autofit_width(string)
            self.assertEqual(got, exp + 7)
