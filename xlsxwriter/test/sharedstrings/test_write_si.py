###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ...sharedstrings import SharedStrings


class TestWriteSi(unittest.TestCase):
    """
    Test the SharedStrings _write_si() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.sharedstrings = SharedStrings()
        self.sharedstrings._set_filehandle(self.fh)

    def test_write_si(self):
        """Test the _write_si() method"""

        self.sharedstrings._write_si('neptune')

        exp = """<si><t>neptune</t></si>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
