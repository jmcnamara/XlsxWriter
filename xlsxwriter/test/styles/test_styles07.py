###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2016, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ..helperfunctions import _xml_to_list
from ...styles import Styles
from ...workbook import Workbook


class TestAssembleStyles(unittest.TestCase):
    """
    Test assembling a complete Styles file.

    """
    def test_assemble_xml_file(self):
        """Test for simple fills."""
        self.maxDiff = None

        fh = StringIO()
        style = Styles()
        style._set_filehandle(fh)

        workbook = Workbook()

        workbook.add_format({'pattern': 1, 'bg_color': 'red'})
        workbook.add_format({'pattern': 11, 'bg_color': 'red'})
        workbook.add_format({'pattern': 11, 'bg_color': 'red', 'fg_color': 'yellow'})
        workbook.add_format({'pattern': 1, 'bg_color': 'red', 'fg_color': 'red'})

        workbook._set_default_xf_indices()
        workbook._prepare_format_properties()

        style._set_style_properties([
            workbook.xf_formats,
            workbook.palette,
            workbook.font_count,
            workbook.num_format_count,
            workbook.border_count,
            workbook.fill_count,
            workbook.custom_colors,
            workbook.dxf_formats,
        ])

        style._assemble_xml_file()
        workbook.fileclosed = 1

        exp = _xml_to_list("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <fonts count="1">
                    <font>
                      <sz val="11"/>
                      <color theme="1"/>
                      <name val="Calibri"/>
                      <family val="2"/>
                      <scheme val="minor"/>
                    </font>
                  </fonts>
                  <fills count="6">
                    <fill>
                      <patternFill patternType="none"/>
                    </fill>
                    <fill>
                      <patternFill patternType="gray125"/>
                    </fill>
                    <fill>
                      <patternFill patternType="solid">
                        <fgColor rgb="FFFF0000"/>
                        <bgColor indexed="64"/>
                      </patternFill>
                    </fill>
                    <fill>
                      <patternFill patternType="lightHorizontal">
                        <bgColor rgb="FFFF0000"/>
                      </patternFill>
                    </fill>
                    <fill>
                      <patternFill patternType="lightHorizontal">
                        <fgColor rgb="FFFFFF00"/>
                        <bgColor rgb="FFFF0000"/>
                      </patternFill>
                    </fill>
                    <fill>
                      <patternFill patternType="solid">
                        <fgColor rgb="FFFF0000"/>
                        <bgColor rgb="FFFF0000"/>
                      </patternFill>
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
                  <cellXfs count="5">
                    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
                    <xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>
                    <xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>
                    <xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/>
                    <xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1"/>
                  </cellXfs>
                  <cellStyles count="1">
                    <cellStyle name="Normal" xfId="0" builtinId="0"/>
                  </cellStyles>
                  <dxfs count="0"/>
                  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
                </styleSheet>
                """)

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)
