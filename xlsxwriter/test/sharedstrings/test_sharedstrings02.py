###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO

from xlsxwriter.sharedstrings import SharedStrings, SharedStringTable

from ..helperfunctions import _xml_to_list

EXPECTED_SHARED_STRINGS_XML = """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
      <si>
        <t>abcdefg</t>
      </si>
      <si>
        <t xml:space="preserve">   abcdefg</t>
      </si>
      <si>
        <t xml:space="preserve">abcdefg   </t>
      </si>
    </sst>
"""


class TestAssembleSharedStrings(unittest.TestCase):
    """
    Test assembling a complete SharedStrings file.

    """

    maxDiff = None

    def setUp(self):
        self.string_table = SharedStringTable()
        self.fh = StringIO()

    def test_shared_string_table_indices(self):

        # Add some strings with leading/trailing whitespace.
        self.assertEqual(self.string_table._get_shared_string_index("abcdefg"), 0)
        self.assertEqual(self.string_table._get_shared_string_index("   abcdefg"), 1)
        self.assertEqual(self.string_table._get_shared_string_index("abcdefg   "), 2)

    def test_assemble_xml_file_with_whitespace_strings(self):
        """Test the _assemble_xml_file() method"""

        self.string_table._get_shared_string_index("abcdefg")
        self.string_table._get_shared_string_index("   abcdefg")
        self.string_table._get_shared_string_index("abcdefg   ")
        self.string_table._sort_string_data()

        sharedstrings = SharedStrings()
        sharedstrings._set_filehandle(self.fh)
        sharedstrings.string_table = self.string_table

        sharedstrings._assemble_xml_file()

        expected = _xml_to_list(EXPECTED_SHARED_STRINGS_XML)
        got = _xml_to_list(self.fh.getvalue())

        self.assertEqual(expected, got)
