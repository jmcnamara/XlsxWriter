###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ..helperfunctions import _xml_to_list
from ...worksheet import Worksheet
from ...sharedstrings import SharedStringTable


class TestAssembleWorksheet(unittest.TestCase):
    """
    Test assembling a complete Worksheet file.

    """
    def test_assemble_xml_file(self):
        """Test writing a worksheet with data out of bounds."""
        self.maxDiff = None

        fh = StringIO()
        worksheet = Worksheet()
        worksheet._set_filehandle(fh)
        worksheet.str_table = SharedStringTable()
        worksheet.select()

        max_row = 1048576
        max_col = 16384
        bound_error = -1

        # Test some out of bound values.
        got = worksheet.write_string(max_row, 0, 'Foo')
        self.assertEqual(got, bound_error)

        got = worksheet.write_string(0, max_col, 'Foo')
        self.assertEqual(got, bound_error)

        got = worksheet.write_string(max_row, max_col, 'Foo')
        self.assertEqual(got, bound_error)

        got = worksheet.write_number(max_row, 0, 123)
        self.assertEqual(got, bound_error)

        got = worksheet.write_number(0, max_col, 123)
        self.assertEqual(got, bound_error)

        got = worksheet.write_number(max_row, max_col, 123)
        self.assertEqual(got, bound_error)

        got = worksheet.write_blank(max_row, 0, None, 'format')
        self.assertEqual(got, bound_error)

        got = worksheet.write_blank(0, max_col, None, 'format')
        self.assertEqual(got, bound_error)

        got = worksheet.write_blank(max_row, max_col, None, 'format')
        self.assertEqual(got, bound_error)

        got = worksheet.write_formula(max_row, 0, '=A1')
        self.assertEqual(got, bound_error)

        got = worksheet.write_formula(0, max_col, '=A1')
        self.assertEqual(got, bound_error)

        got = worksheet.write_formula(max_row, max_col, '=A1')
        self.assertEqual(got, bound_error)

        got = worksheet.write_array_formula(0, 0, 0, max_col, '=A1')
        self.assertEqual(got, bound_error)

        got = worksheet.write_array_formula(0, 0, max_row, 0, '=A1')
        self.assertEqual(got, bound_error)

        got = worksheet.write_array_formula(0, max_col, 0, 0, '=A1')
        self.assertEqual(got, bound_error)

        got = worksheet.write_array_formula(max_row, 0, 0, 0, '=A1')
        self.assertEqual(got, bound_error)

        got = worksheet.write_array_formula(max_row, max_col, max_row, max_col, '=A1')
        self.assertEqual(got, bound_error)

        # Column out of bounds.
        got = worksheet.set_column(6, max_col, 17)
        self.assertEqual(got, bound_error)

        got = worksheet.set_column(max_col, 6, 17)
        self.assertEqual(got, bound_error)

        # Row out of bounds.
        worksheet.set_row(max_row, 30)

        # Reverse man and min column numbers
        worksheet.set_column(0, 3, 17)

        # Write some valid strings.
        worksheet.write_string(0, 0, 'Foo')
        worksheet.write_string(2, 0, 'Bar')
        worksheet.write_string(2, 3, 'Baz')

        worksheet._assemble_xml_file()

        exp = _xml_to_list("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <dimension ref="A1:D3"/>
                  <sheetViews>
                    <sheetView tabSelected="1" workbookViewId="0"/>
                  </sheetViews>
                  <sheetFormatPr defaultRowHeight="15"/>
                  <cols>
                    <col min="1" max="4" width="17.7109375" customWidth="1"/>
                  </cols>
                  <sheetData>
                    <row r="1" spans="1:4">
                      <c r="A1" t="s">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="3" spans="1:4">
                      <c r="A3" t="s">
                        <v>1</v>
                      </c>
                      <c r="D3" t="s">
                        <v>2</v>
                      </c>
                    </row>
                  </sheetData>
                  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
                </worksheet>
                """)

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)
