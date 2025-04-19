###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO

from xlsxwriter.sharedstrings import SharedStrings, SharedStringTable


class TestWriteSst(unittest.TestCase):
    """
    Test the SharedStrings _write_sst() method.

    """

    def setUp(self):
        self.fh = StringIO()
        self.sharedstrings = SharedStrings()
        self.sharedstrings._set_filehandle(self.fh)

    def test_write_sst(self):
        """Test the _write_sst() method"""

        string_table = SharedStringTable()

        # Add some strings and check the returned indices.
        string_table._get_shared_string_index("neptune")
        string_table._get_shared_string_index("neptune")
        string_table._get_shared_string_index("neptune")
        string_table._get_shared_string_index("mars")
        string_table._get_shared_string_index("venus")
        string_table._get_shared_string_index("mars")
        string_table._get_shared_string_index("venus")
        self.sharedstrings.string_table = string_table

        self.sharedstrings._write_sst()

        exp = """<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="7" uniqueCount="3">"""
        got = self.fh.getvalue()

        self.assertEqual(exp, got)
