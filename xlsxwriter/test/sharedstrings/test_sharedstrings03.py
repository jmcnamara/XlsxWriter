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


class TestAssembleSharedStrings(unittest.TestCase):
    """
    Test assembling a complete SharedStrings file.

    """

    def test_assemble_xml_file(self):
        """Test a plain string that matches the XML of a rich string."""

        string_table = SharedStringTable()

        index = string_table._get_rich_string_index("<r><t>abc</t></r>")
        self.assertEqual(index, 0)

        index = string_table._get_shared_string_index("<r><t>abc</t></r>")
        self.assertEqual(index, 1)

        string_table._sort_string_data()

        fh = StringIO()
        sharedstrings = SharedStrings()
        sharedstrings._set_filehandle(fh)
        sharedstrings.string_table = string_table

        sharedstrings._assemble_xml_file()

        exp = _xml_to_list("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
                  <si>
                    <r>
                      <t>abc</t>
                    </r>
                  </si>
                  <si>
                    <t>&lt;r&gt;&lt;t&gt;abc&lt;/t&gt;&lt;/r&gt;</t>
                  </si>
                </sst>
                """)

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(exp, got)
