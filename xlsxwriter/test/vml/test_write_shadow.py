###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...vml import Vml


class TestWriteVshadow(unittest.TestCase):
    """
    Test the Vml _write_shadow() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.vml = Vml()
        self.vml._set_filehandle(self.fh)

    def test_write_shadow(self):
        """Test the _write_shadow() method"""

        self.vml._write_shadow()

        exp = """<v:shadow on="t" color="black" obscured="t"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
