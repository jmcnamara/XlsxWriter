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

from xlsxwriter.sharedstrings import SharedStrings


class TestInitialisation(unittest.TestCase):
    """
    Test initialisation of the SharedStrings class and call a method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.sharedstrings = SharedStrings()
        self.sharedstrings._set_filehandle(self.fh)

    def test_xml_declaration(self):
        """Test Sharedstrings xml_declaration()"""

        self.sharedstrings._xml_declaration()

        exp = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n"""
        got = self.fh.getvalue()

        self.assertEqual(exp, got)
