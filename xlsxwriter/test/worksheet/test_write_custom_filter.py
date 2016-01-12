###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...worksheet import Worksheet


class TestWriteCustomFilter(unittest.TestCase):
    """
    Test the Worksheet _write_custom_filter() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_write_custom_filter(self):
        """Test the _write_custom_filter() method"""

        self.worksheet._write_custom_filter(4, 3000)

        exp = """<customFilter operator="greaterThan" val="3000"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
