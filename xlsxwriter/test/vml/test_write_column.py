###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...vml import Vml


class TestWriteXColumn(unittest.TestCase):
    """
    Test the Vml _write_column() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.vml = Vml()
        self.vml._set_filehandle(self.fh)

    def test_write_column(self):
        """Test the _write_column() method"""

        self.vml._write_column(2)

        exp = """<x:Column>2</x:Column>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
