###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...vml import Vml


class TestWriteOshapelayout(unittest.TestCase):
    """
    Test the Vml _write_shapelayout() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.vml = Vml()
        self.vml._set_filehandle(self.fh)

    def test_write_shapelayout(self):
        """Test the _write_shapelayout() method"""

        self.vml._write_shapelayout(1)

        exp = """<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
