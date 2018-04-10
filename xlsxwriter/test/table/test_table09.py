###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ..helperfunctions import _xml_to_list
from ...table import Table
from ...worksheet import Worksheet
from ...workbook import WorksheetMeta
from ...sharedstrings import SharedStringTable


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

        # Set the table properties.
        worksheet.add_table('B2:K8', {'total_row': 1,
                                      'columns': [{'total_string': 'Total'},
                                                  {},
                                                  {'total_function': 'Average'},
                                                  {'total_function': 'COUNT'},
                                                  {'total_function': 'count_nums'},
                                                  {'total_function': 'max'},
                                                  {'total_function': 'min'},
                                                  {'total_function': 'sum'},
                                                  {'total_function': 'std Dev'},
                                                  {'total_function': 'var'}
                                                  ]})
        worksheet._prepare_tables(1, {})

        fh = StringIO()
        table = Table()
        table._set_filehandle(fh)

        table._set_properties(worksheet.tables[0])

        table._assemble_xml_file()

        exp = _xml_to_list("""
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
                """)

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)
