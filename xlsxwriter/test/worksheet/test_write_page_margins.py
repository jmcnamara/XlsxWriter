###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013, John McNamara, jmcnamara@cpan.org
#

import unittest
from StringIO import StringIO
from ...worksheet import Worksheet


class TestWritePageMargins(unittest.TestCase):
    """
    Test the Worksheet _write_page_margins() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet(self.fh)

    def test_write_page_margins(self):
        """Test the _write_page_margins() method"""

        self.worksheet._write_page_margins()

        exp = """<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)


if __name__ == '__main__':
    unittest.main()
