###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...vml import Vml


class TestWriteXMoveWithCells(unittest.TestCase):
    """
    Test the Vml _write_move_with_cells() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.vml = Vml()
        self.vml._set_filehandle(self.fh)

    def test_write_move_with_cells(self):
        """Test the _write_move_with_cells() method"""

        self.vml._write_move_with_cells()

        exp = """<x:MoveWithCells/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
