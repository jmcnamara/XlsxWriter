###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...drawing import Drawing


class TestWriteAgraphicFrameLocks(unittest.TestCase):
    """
    Test the Drawing _write_a_graphic_frame_locks() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.drawing = Drawing()
        self.drawing._set_filehandle(self.fh)

    def test_write_a_graphic_frame_locks(self):
        """Test the _write_a_graphic_frame_locks() method"""

        self.drawing._write_a_graphic_frame_locks()

        exp = """<a:graphicFrameLocks noGrp="1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
