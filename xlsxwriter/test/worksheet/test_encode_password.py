###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO

from xlsxwriter.worksheet import Worksheet


class TestEncodePassword(unittest.TestCase):
    """
    Test the Worksheet _encode_password() methods.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)
        self.encode = self.worksheet._encode_password

    def test_encode_password_text(self):
        self.assertEqual(self.encode("password"), "83AF")
        self.assertEqual(self.encode("This is a longer phrase"), "D14E")

    def test_encode_password_short_numeric(self):
        self.assertEqual(self.encode("0"), "CE2A")
        self.assertEqual(self.encode("01"), "CEED")
        self.assertEqual(self.encode("012"), "CF7C")
        self.assertEqual(self.encode("0123"), "CC4B")
        self.assertEqual(self.encode("01234"), "CACA")
        self.assertEqual(self.encode("012345"), "C789")
        self.assertEqual(self.encode("0123456"), "DC88")
        self.assertEqual(self.encode("01234567"), "EB87")
        self.assertEqual(self.encode("012345678"), "9B86")

    def test_encode_password_medium_numeric(self):
        self.assertEqual(self.encode("0123456789"), "FF84")
        self.assertEqual(self.encode("01234567890"), "FF86")
        self.assertEqual(self.encode("012345678901"), "EF87")
        self.assertEqual(self.encode("0123456789012"), "AF8A")
        self.assertEqual(self.encode("01234567890123"), "EF90")
        self.assertEqual(self.encode("012345678901234"), "EFA5")
        self.assertEqual(self.encode("0123456789012345"), "EFD0")

    def test_encode_password_long_numeric(self):
        self.assertEqual(self.encode("01234567890123456"), "EF09")
        self.assertEqual(self.encode("012345678901234567"), "EEB2")
        self.assertEqual(self.encode("0123456789012345678"), "ED33")
        self.assertEqual(self.encode("01234567890123456789"), "EA14")
        self.assertEqual(self.encode("012345678901234567890"), "E615")
        self.assertEqual(self.encode("0123456789012345678901"), "FE96")
        self.assertEqual(self.encode("01234567890123456789012"), "CC97")
        self.assertEqual(self.encode("012345678901234567890123"), "AA98")
        self.assertEqual(self.encode("0123456789012345678901234"), "FA98")
        self.assertEqual(self.encode("01234567890123456789012345"), "D298")
        self.assertEqual(self.encode("0123456789012345678901234567890"), "D2D3")
