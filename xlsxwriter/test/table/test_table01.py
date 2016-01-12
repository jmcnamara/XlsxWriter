###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
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

        worksheet.add_table('C3:F13')
        worksheet._prepare_tables(1, {})

        fh = StringIO()
        table = Table()
        table._set_filehandle(fh)

        table._set_properties(worksheet.tables[0])

        table._assemble_xml_file()

        exp = _xml_to_list("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C3:F13" totalsRowShown="0">
                  <autoFilter ref="C3:F13"/>
                  <tableColumns count="4">
                    <tableColumn id="1" name="Column1"/>
                    <tableColumn id="2" name="Column2"/>
                    <tableColumn id="3" name="Column3"/>
                    <tableColumn id="4" name="Column4"/>
                  </tableColumns>
                  <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
                """)

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)
