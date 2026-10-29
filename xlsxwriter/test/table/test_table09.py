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

from xlsxwriter.sharedstrings import SharedStringTable
from xlsxwriter.table import Table
from xlsxwriter.workbook import WorksheetMeta
from xlsxwriter.worksheet import Worksheet

from ..helperfunctions import _xml_to_list

EXPECTED_TABLE_XML = """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="B2:K8" totalsRowCount="1">
      <autoFilter ref="B2:K7"/>
      <tableColumns count="10">
        <tableColumn id="1" name="Column1" totalsRowLabel="Total"/>
        <tableColumn id="2" name="Column2"/>
        <tableColumn id="3" name="Column3" totalsRowFunction="average"/>
        <tableColumn id="4" name="Column4" totalsRowFunction="count"/>
        <tableColumn id="5" name="Column5" totalsRowFunction="countNums"/>
        <tableColumn id="6" name="Column6" totalsRowFunction="max"/>
        <tableColumn id="7" name="Column7" totalsRowFunction="min"/>
        <tableColumn id="8" name="Column8" totalsRowFunction="sum"/>
        <tableColumn id="9" name="Column9" totalsRowFunction="stdDev"/>
        <tableColumn id="10" name="Column10" totalsRowFunction="var"/>
      </tableColumns>
      <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
    </table>
"""


class TestAssembleTable(unittest.TestCase):
    """
    Test assembling a complete Table file.

    """

    maxDiff = None

    def setUp(self):
        self.worksheet = Worksheet()
        self.worksheet.worksheet_meta = WorksheetMeta()
        self.worksheet.str_table = SharedStringTable()

        self.fh = StringIO()
        self.table = Table()
        self.table._set_filehandle(self.fh)

    def _assemble_and_get_xml(self):
        self.table._set_properties(self.worksheet.tables[0])
        self.table._assemble_xml_file()
        return _xml_to_list(self.fh.getvalue())

    def test_assemble_xml_file(self):
        """Test writing a table"""
        # Set the table properties.
        self.worksheet.add_table(
            "B2:K8",
            {
                "total_row": 1,
                "columns": [
                    {"total_string": "Total"},
                    {},
                    {"total_function": "average"},
                    {"total_function": "count"},
                    {"total_function": "count_nums"},
                    {"total_function": "max"},
                    {"total_function": "min"},
                    {"total_function": "sum"},
                    {"total_function": "stdDev"},
                    {"total_function": "var"},
                ],
            },
        )
        self.worksheet._prepare_tables(1, {})

        exp = _xml_to_list(EXPECTED_TABLE_XML)
        got = self._assemble_and_get_xml()

        self.assertEqual(exp, got)
