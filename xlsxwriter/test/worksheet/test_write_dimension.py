###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ...worksheet import Worksheet


class TestWriteDimension(unittest.TestCase):
    """
    Test the Worksheet _write_dimension() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.worksheet = Worksheet()
        self.worksheet._set_filehandle(self.fh)

    def test_write_dimension(self):
        """Test the _write_dimension() method"""

        self.worksheet._write_dimension()

        exp = """<dimension ref="A1"/>"""
        got = self.fh.getvalue()

        self.assertEqual(got, exp)
