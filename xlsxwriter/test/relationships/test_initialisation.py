###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...relationships import Relationships


class TestInitialisation(unittest.TestCase):
    """
    Test initialisation of the Relationships class and call a method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.relationships = Relationships()
        self.relationships._set_filehandle(self.fh)

    def test_xml_declaration(self):
        """Test Relationships xml_declaration()"""

        self.relationships._xml_declaration()

        exp = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
