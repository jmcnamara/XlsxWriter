###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...styles import Styles


class TestWriteDefaultFill(unittest.TestCase):
    """
    Test the Styles _write_default_fill() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.styles = Styles()
        self.styles._set_filehandle(self.fh)

    def test_write_default_fill(self):
        """Test the _write_default_fill() method"""

        self.styles._write_default_fill('none')

        exp = """<fill><patternFill patternType="none"/></fill>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
