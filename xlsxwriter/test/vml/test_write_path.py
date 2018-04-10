###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...vml import Vml


class TestWriteVpath(unittest.TestCase):
    """
    Test the Vml _write_path() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.vml = Vml()
        self.vml._set_filehandle(self.fh)

    def test_write_comment_path_1(self):
        """Test the _write_comment_path() method"""

        self.vml._write_comment_path('t', 'rect')

        exp = """<v:path gradientshapeok="t" o:connecttype="rect"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_comment_path_2(self):
        """Test the _write_comment_path() method"""

        self.vml._write_comment_path(None, 'none')

        exp = """<v:path o:connecttype="none"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_button_path(self):
        """Test the _write_button_path() method"""

        self.vml._write_button_path()

        exp = """<v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
