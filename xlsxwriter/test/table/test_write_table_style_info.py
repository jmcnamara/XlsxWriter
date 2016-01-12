###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...table import Table


class TestWriteTableStyleInfo(unittest.TestCase):
    """
    Test the Table _write_table_style_info() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.table = Table()
        self.table._set_filehandle(self.fh)

    def test_write_table_style_info(self):
        """Test the _write_table_style_info() method"""

        self.table.properties = {
            'style': 'TableStyleMedium9',
            'show_first_col': 0,
            'show_last_col': 0,
            'show_row_stripes': 1,
            'show_col_stripes': 0,
        }

        self.table._write_table_style_info()

        exp = """<tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
