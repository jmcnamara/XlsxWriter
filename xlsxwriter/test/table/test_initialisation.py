###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...table import Table


class TestInitialisation(unittest.TestCase):
    """
    Test initialisation of the Table class and call a method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.table = Table()
        self.table._set_filehandle(self.fh)

    def test_xml_declaration(self):
        """Test Table xml_declaration()"""

        self.table._xml_declaration()

        exp = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
