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

from xlsxwriter.sharedstrings import SharedStringTable
from xlsxwriter.table import Table
from xlsxwriter.workbook import WorksheetMeta
from xlsxwriter.worksheet import Worksheet

from ..helperfunctions import _xml_to_list


class TestAssembleTable(unittest.TestCase):
    """
    Test assembling a complete Table file.

    """

    def test_assemble_xml_file(self):
        """Test writing a table"""
        self.maxDiff = None

        worksheet = Worksheet()
        worksheet.worksheet_meta = WorksheetMeta()
        worksheet.str_table = SharedStringTable()

        worksheet.add_table("D4:I15", {"style": "Table Style Light 17"})
        worksheet._prepare_tables(1, {})

        fh = StringIO()
        table = Table()
        table._set_filehandle(fh)

        table._set_properties(worksheet.tables[0])

        table._assemble_xml_file()

        exp = _xml_to_list(
            """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="D4:I15" totalsRowShown="0">
                  <autoFilter ref="D4:I15"/>
                  <tableColumns count="6">
                    <tableColumn id="1" name="Column1"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3"/>
                    <tableColumn id="4" name="Column4"/>
                    <tableColumn id="5" name="Column5"/>
                    <tableColumn id="6" name="Column6"/>
                  </tableColumns>
                  <tableStyleInfo name="TableStyleLight17" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
                """
        )

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(exp, got)
