###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...format import Format


class TestInitialisation(unittest.TestCase):
    """
    Test initialisation of the Format class and call a method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.format = Format()
        self.format._set_filehandle(self.fh)

    def test_xml_declaration(self):
        """Test Format xml_declaration()"""

        self.format._xml_declaration()

        exp = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
