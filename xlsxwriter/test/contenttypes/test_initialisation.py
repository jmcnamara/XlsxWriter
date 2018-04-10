###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...contenttypes import ContentTypes


class TestInitialisation(unittest.TestCase):
    """
    Test initialisation of the ContentTypes class and call a method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.contenttypes = ContentTypes()
        self.contenttypes._set_filehandle(self.fh)

    def test_xml_declaration(self):
        """Test ContentTypes xml_declaration()"""

        self.contenttypes._xml_declaration()

        exp = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
