###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ...vml import Vml


class TestWriteVshapetype(unittest.TestCase):
    """
    Test the Vml _write_shapetype() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.vml = Vml()
        self.vml._set_filehandle(self.fh)

    def test_write_comment_shapetype(self):
        """Test the _write_comment_shapetype() method"""

        self.vml._write_comment_shapetype()

        exp = """<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)

    def test_write_button_shapetype(self):
        """Test the _write_button_shapetype() method"""

        self.vml._write_button_shapetype()

        exp = """<v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/><o:lock v:ext="edit" shapetype="t"/></v:shapetype>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
