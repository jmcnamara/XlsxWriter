###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...worksheet import Worksheet


class TestWriteCustomFilters(unittest.TestCase):
    """
    Test the Worksheet _write_custom_filters() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_write_custom_filters_1(self):
        """Test the _write_custom_filters() method"""

        self.worksheet._write_custom_filters([4, 4000])

        exp = """<customFilters><customFilter operator="greaterThan" val="4000"/></customFilters>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_custom_filters_2(self):
        """Test the _write_custom_filters() method"""

        self.worksheet._write_custom_filters([4, 3000, 0, 1, 8000])

        exp = """<customFilters and="1"><customFilter operator="greaterThan" val="3000"/><customFilter operator="lessThan" val="8000"/></customFilters>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
