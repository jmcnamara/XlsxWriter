###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...vml import Vml


class TestWriteDiv(unittest.TestCase):
    """
    Test the Vml _write_div() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.vml = Vml()
        self.vml._set_filehandle(self.fh)

    def test_write_div(self):
        """Test the _write_div() method"""

        self.vml._write_div('left')

        exp = """<div style="text-align:left"></div>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
