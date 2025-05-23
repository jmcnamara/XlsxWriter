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

from xlsxwriter.styles import Styles
from xlsxwriter.workbook import Workbook

from ..helperfunctions import _xml_to_list


class TestAssembleStyles(unittest.TestCase):
    """
    Test assembling a complete Styles file.

    """

    def test_assemble_xml_file(self):
        """Test for simple font styles."""
        self.maxDiff = None

        fh = StringIO()
        style = Styles()
        style._set_filehandle(fh)

        workbook = Workbook()

        workbook.add_format({"bold": 1})
        workbook.add_format({"italic": 1})
        workbook.add_format({"bold": 1, "italic": 1})

        workbook._set_default_xf_indices()
        workbook._prepare_format_properties()

        style._set_style_properties(
            [
                workbook.xf_formats,
                workbook.palette,
                workbook.font_count,
                workbook.num_formats,
                workbook.border_count,
                workbook.fill_count,
                workbook.custom_colors,
                workbook.dxf_formats,
                workbook.has_comments,
            ]
        )

        style._assemble_xml_file()
        workbook.fileclosed = 1

        exp = _xml_to_list(
            """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <fonts count="4">
                    <font>
                      <sz val="11"/>
                      <color theme="1"/>
                      <name val="Calibri"/>
                      <family val="2"/>
                      <scheme val="minor"/>
                    </font>
                    <font>
                      <b/>
                      <sz val="11"/>
                      <color theme="1"/>
                      <name val="Calibri"/>
                      <family val="2"/>
                      <scheme val="minor"/>
                    </font>
                    <font>
                      <i/>
                      <sz val="11"/>
                      <color theme="1"/>
                      <name val="Calibri"/>
                      <family val="2"/>
                      <scheme val="minor"/>
                    </font>
                    <font>
                      <b/>
                      <i/>
                      <sz val="11"/>
                      <color theme="1"/>
                      <name val="Calibri"/>
                      <family val="2"/>
                      <scheme val="minor"/>
                    </font>
                  </fonts>
                  <fills count="2">
                    <fill>
                      <patternFill patternType="none"/>
                    </fill>
                    <fill>
                      <patternFill patternType="gray125"/>
                    </fill>
                  </fills>
                  <borders count="1">
                    <border>
                      <left/>
                      <right/>
                      <top/>
                      <bottom/>
                      <diagonal/>
                    </border>
                  </borders>
                  <cellStyleXfs count="1">
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
                  </cellStyleXfs>
                  <cellXfs count="4">
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
                    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
                    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
                    <xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
                  </cellXfs>
                  <cellStyles count="1">
                    <cellStyle name="Normal" xfId="0" builtinId="0"/>
                  </cellStyles>
                  <dxfs count="0"/>
                  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
                </styleSheet>
                """
        )

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(exp, got)
