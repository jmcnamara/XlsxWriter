###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...vml import Vml


class TestWriteXRow(unittest.TestCase):
    """
    Test the Vml _write_row() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.vml = Vml()
        self.vml._set_filehandle(self.fh)

    def test_write_row(self):
        """Test the _write_row() method"""

        self.vml._write_row(2)

        exp = """<x:Row>2</x:Row>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
