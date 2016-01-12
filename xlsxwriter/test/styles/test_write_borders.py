###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...styles import Styles
from ...format import Format


class TestWriteBorders(unittest.TestCase):
    """
    Test the Styles _write_borders() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.styles = Styles()
        self.styles._set_filehandle(self.fh)

    def test_write_borders(self):
        """Test the _write_borders() method"""

        xf_format = Format()
        xf_format.has_border = 1

        self.styles._set_style_properties([[xf_format], None, 0, 0, 1, 0, [], []])

        self.styles._write_borders()

        exp = """<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
