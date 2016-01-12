###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...vml import Vml


class TestWriteOidmap(unittest.TestCase):
    """
    Test the Vml _write_idmap() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.vml = Vml()
        self.vml._set_filehandle(self.fh)

    def test_write_idmap(self):
        """Test the _write_idmap() method"""

        self.vml._write_idmap(1)

        exp = """<o:idmap v:ext="edit" data="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
