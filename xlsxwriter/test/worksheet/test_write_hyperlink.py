###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...worksheet import Worksheet


class TestWriteHyperlink(unittest.TestCase):
    """
    Test the Worksheet _write_hyperlink() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_write_hyperlink_external(self):
        """Test the _write_hyperlink() method"""

        self.worksheet._write_hyperlink_external(0, 0, 1)

        exp = """<hyperlink ref="A1" r:id="rId1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_hyperlink_internal_1(self):
        """Test the _write_hyperlink() method"""

        self.worksheet._write_hyperlink_internal(0, 0, 'Sheet2!A1', 'Sheet2!A1')

        exp = """<hyperlink ref="A1" location="Sheet2!A1" display="Sheet2!A1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_hyperlink_internal_2(self):
        """Test the _write_hyperlink() method"""

        self.worksheet._write_hyperlink_internal(4, 0, "'Data Sheet'!D5", "'Data Sheet'!D5")

        exp = """<hyperlink ref="A5" location="'Data Sheet'!D5" display="'Data Sheet'!D5"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_hyperlink_internal_3(self):
        """Test the _write_hyperlink() method"""

        self.worksheet._write_hyperlink_internal(17, 0, 'Sheet2!A1', 'Sheet2!A1', 'Screen Tip 1')

        exp = """<hyperlink ref="A18" location="Sheet2!A1" tooltip="Screen Tip 1" display="Sheet2!A1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
