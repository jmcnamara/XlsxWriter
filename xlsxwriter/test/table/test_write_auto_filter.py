###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...table import Table


class TestWriteAutoFilter(unittest.TestCase):
    """
    Test the Table _write_auto_filter() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.table = Table()
        self.table._set_filehandle(self.fh)

    def test_write_auto_filter(self):
        """Test the _write_auto_filter() method"""

        self.table.properties['autofilter'] = 'C3:F13'

        self.table._write_auto_filter()

        exp = """<autoFilter ref="C3:F13"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
