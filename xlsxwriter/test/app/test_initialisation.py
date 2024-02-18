###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ...app import App


class TestInitialisation(unittest.TestCase):
    """
    Test initialisation of the App class and call a method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.app = App()
        self.app._set_filehandle(self.fh)

    def test_xml_declaration(self):
        """Test App xml_declaration()"""

        self.app._xml_declaration()

        exp = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
