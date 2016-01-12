###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...styles import Styles


class TestWriteColors(unittest.TestCase):
    """
    Test the Styles _write_colors() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.styles = Styles()
        self.styles._set_filehandle(self.fh)

    def test_write_colors1(self):
        """Test the _write_colors() method"""

        self.styles.custom_colors = ['FF26DA55']
        self.styles._write_colors()

        exp = """<colors><mruColors><color rgb="FF26DA55"/></mruColors></colors>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_colors2(self):
        """Test the _write_colors() method"""

        self.styles.custom_colors = ['FF26DA55', 'FF792DC8', 'FF646462']
        self.styles._write_colors()

        exp = """<colors><mruColors><color rgb="FF646462"/><color rgb="FF792DC8"/><color rgb="FF26DA55"/></mruColors></colors>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_colors3(self):
        """Test the _write_colors() method"""

        self.styles.custom_colors = ['FF792DC8', 'FF646462', 'FF5EA29C',
                                     'FF583AC6', 'FFE31DAF', 'FFA1A759',
                                     'FF600FF1', 'FF0CF49C', 'FFE3FA06',
                                     'FF913AC6', 'FFB97847', 'FFD97827']

        self.styles._write_colors()

        exp = """<colors><mruColors><color rgb="FFD97827"/><color rgb="FFB97847"/><color rgb="FF913AC6"/><color rgb="FFE3FA06"/><color rgb="FF0CF49C"/><color rgb="FF600FF1"/><color rgb="FFA1A759"/><color rgb="FFE31DAF"/><color rgb="FF583AC6"/><color rgb="FF5EA29C"/></mruColors></colors>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
