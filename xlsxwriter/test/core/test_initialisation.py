###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ...core import Core


class TestInitialisation(unittest.TestCase):
    """
    Test initialisation of the Core class and call a method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.core = Core()
        self.core._set_filehandle(self.fh)

    def test_xml_declaration(self):
        """Test Core xml_declaration()"""

        self.core._xml_declaration()

        exp = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
