###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013, John McNamara, jmcnamara@cpan.org
#

import unittest
from StringIO import StringIO
from collections import namedtuple
from ...worksheet import Worksheet


class TestWriteCell(unittest.TestCase):
    """
    Test the Worksheet _write_cell() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_write_cell(self):
        """Test the _write_cell() method"""

        cell_tuple = namedtuple('Number', 'number, format')
        cell = cell_tuple(1, None)

        self.worksheet._write_cell(0, 0, cell)

        exp = """<c r="A1"><v>1</v></c>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)


if __name__ == '__main__':
    unittest.main()
