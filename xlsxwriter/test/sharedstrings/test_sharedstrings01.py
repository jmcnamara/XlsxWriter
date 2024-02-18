###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ..helperfunctions import _xml_to_list
from ...sharedstrings import SharedStringTable
from ...sharedstrings import SharedStrings


class TestAssembleSharedStrings(unittest.TestCase):
    """
    Test assembling a complete SharedStrings file.

    """

    def test_assemble_xml_file(self):
        """Test the _write_sheet_data() method"""

        string_table = SharedStringTable()

        # Add some strings and check the returned indices.
        index = string_table._get_shared_string_index("neptune")
        self.assertEqual(index, 0)

        index = string_table._get_shared_string_index("neptune")
        self.assertEqual(index, 0)

        index = string_table._get_shared_string_index("neptune")
        self.assertEqual(index, 0)

        index = string_table._get_shared_string_index("mars")
        self.assertEqual(index, 1)

        index = string_table._get_shared_string_index("venus")
        self.assertEqual(index, 2)

        index = string_table._get_shared_string_index("mars")
        self.assertEqual(index, 1)

        index = string_table._get_shared_string_index("venus")
        self.assertEqual(index, 2)

        string_table._sort_string_data()

        fh = StringIO()
        sharedstrings = SharedStrings()
        sharedstrings._set_filehandle(fh)
        sharedstrings.string_table = string_table

        sharedstrings._assemble_xml_file()

        exp = _xml_to_list(
            """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="7" uniqueCount="3">
                  <si>
                    <t>neptune</t>
                  </si>
                  <si>
                    <t>mars</t>
                  </si>
                  <si>
                    <t>venus</t>
                  </si>
                </sst>
                """
        )

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)
