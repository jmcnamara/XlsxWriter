###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...table import Table


class TestWriteTableColumn(unittest.TestCase):
    """
    Test the Table _write_table_column() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.table = Table()
        self.table._set_filehandle(self.fh)

    def test_write_table_column(self):
        """Test the _write_table_column() method"""

        self.table._write_table_column({'name': 'Column1', 'id': 1})

        exp = """<tableColumn id="1" name="Column1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
