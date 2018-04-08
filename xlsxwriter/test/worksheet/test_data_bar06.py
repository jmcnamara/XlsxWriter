###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2018, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...compatibility import StringIO
from ..helperfunctions import _xml_to_list
from ...worksheet import Worksheet


class TestAssembleWorksheet(unittest.TestCase):
    """
    Test assembling a complete Worksheet file.

    """
    def test_assemble_xml_file(self):
        """Test writing a worksheet with conditional formatting."""
        self.maxDiff = None

        fh = StringIO()
        worksheet = Worksheet()
        worksheet._set_filehandle(fh)
        worksheet.select()
        worksheet.index = 0

        worksheet.conditional_format('A1',
                                     {'type': 'data_bar',
                                      'bar_negative_color_same': True,
                                      })

        worksheet.conditional_format('A2:B2',
                                     {'type': 'data_bar',
                                      'bar_color': '#63C384',
                                      'bar_negative_border_color': '#92D050',
                                      })

        worksheet.conditional_format('A3:C3',
                                     {'type': 'data_bar',
                                      'bar_color': '#FF555A',
                                      'bar_negative_border_color_same': True,
                                      })

        worksheet._assemble_xml_file()

        exp = _xml_to_list("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
                  <dimension ref="A1"/>
                  <sheetViews>
                    <sheetView tabSelected="1" workbookViewId="0"/>
                  </sheetViews>
                  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
                  <sheetData/>
                  <conditionalFormatting sqref="A1">
                    <cfRule type="dataBar" priority="1">
                      <dataBar>
                        <cfvo type="min"/>
                        <cfvo type="max"/>
                        <color rgb="FF638EC6"/>
                      </dataBar>
                      <extLst>
                        <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                          <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000001}</x14:id>
                        </ext>
                      </extLst>
                    </cfRule>
                  </conditionalFormatting>
                  <conditionalFormatting sqref="A2:B2">
                    <cfRule type="dataBar" priority="2">
                      <dataBar>
                        <cfvo type="min"/>
                        <cfvo type="max"/>
                        <color rgb="FF63C384"/>
                      </dataBar>
                      <extLst>
                        <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                          <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000002}</x14:id>
                        </ext>
                      </extLst>
                    </cfRule>
                  </conditionalFormatting>
                  <conditionalFormatting sqref="A3:C3">
                    <cfRule type="dataBar" priority="3">
                      <dataBar>
                        <cfvo type="min"/>
                        <cfvo type="max"/>
                        <color rgb="FFFF555A"/>
                      </dataBar>
                      <extLst>
                        <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
                          <x14:id>{DA7ABA51-AAAA-BBBB-0001-000000000003}</x14:id>
                        </ext>
                      </extLst>
                    </cfRule>
                  </conditionalFormatting>
                  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
                      <x14:conditionalFormattings>
                        <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                          <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
                            <x14:dataBar minLength="0" maxLength="100" border="1" negativeBarColorSameAsPositive="1" negativeBarBorderColorSameAsPositive="0">
                              <x14:cfvo type="autoMin"/>
                              <x14:cfvo type="autoMax"/>
                              <x14:borderColor rgb="FF638EC6"/>
                              <x14:negativeBorderColor rgb="FFFF0000"/>
                              <x14:axisColor rgb="FF000000"/>
                            </x14:dataBar>
                          </x14:cfRule>
                          <xm:sqref>A1</xm:sqref>
                        </x14:conditionalFormatting>
                        <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                          <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000002}">
                            <x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0">
                              <x14:cfvo type="autoMin"/>
                              <x14:cfvo type="autoMax"/>
                              <x14:borderColor rgb="FF63C384"/>
                              <x14:negativeFillColor rgb="FFFF0000"/>
                              <x14:negativeBorderColor rgb="FF92D050"/>
                              <x14:axisColor rgb="FF000000"/>
                            </x14:dataBar>
                          </x14:cfRule>
                          <xm:sqref>A2:B2</xm:sqref>
                        </x14:conditionalFormatting>
                        <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                          <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000003}">
                            <x14:dataBar minLength="0" maxLength="100" border="1">
                              <x14:cfvo type="autoMin"/>
                              <x14:cfvo type="autoMax"/>
                              <x14:borderColor rgb="FFFF555A"/>
                              <x14:negativeFillColor rgb="FFFF0000"/>
                              <x14:axisColor rgb="FF000000"/>
                            </x14:dataBar>
                          </x14:cfRule>
                          <xm:sqref>A3:C3</xm:sqref>
                        </x14:conditionalFormatting>
                      </x14:conditionalFormattings>
                    </ext>
                  </extLst>
                </worksheet>
                """)

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)
