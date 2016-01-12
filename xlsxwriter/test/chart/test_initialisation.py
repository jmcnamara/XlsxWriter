###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...chart import Chart


class TestInitialisation(unittest.TestCase):
    """
    Test initialisation of the Chart class and call a method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.chart = Chart()
        self.chart._set_filehandle(self.fh)

    def test_xml_declaration(self):
        """Test Chart xml_declaration()"""

        self.chart._xml_declaration()

        exp = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
