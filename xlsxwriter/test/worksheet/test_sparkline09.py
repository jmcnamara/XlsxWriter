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
        """Test writing a worksheet with no cell data."""
        self.maxDiff = None

        fh = StringIO()
        worksheet = Worksheet()
        worksheet._set_filehandle(fh)
        worksheet.select()
        worksheet.name = 'Sheet1'
        worksheet.excel_version = 2010

        data = [-2, 2, 3, -1, 0]
        worksheet.write_row('A1', data)

        # Set up sparklines.

        # Test all the styles.
        for i in range(36):
            row = i + 1
            sparkrange = 'Sheet1!A%d:E%d' % (row, row)
            worksheet.write_row(i, 0, data)
            worksheet.add_sparkline(i, 5, {'range': sparkrange,
                                           'style': row})

        worksheet._assemble_xml_file()

        exp = _xml_to_list("""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
                  <dimension ref="A1:E36"/>
                  <sheetViews>
                    <sheetView tabSelected="1" workbookViewId="0"/>
                  </sheetViews>
                  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
                  <sheetData>
                    <row r="1" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A1">
                        <v>-2</v>
                      </c>
                      <c r="B1">
                        <v>2</v>
                      </c>
                      <c r="C1">
                        <v>3</v>
                      </c>
                      <c r="D1">
                        <v>-1</v>
                      </c>
                      <c r="E1">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="2" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A2">
                        <v>-2</v>
                      </c>
                      <c r="B2">
                        <v>2</v>
                      </c>
                      <c r="C2">
                        <v>3</v>
                      </c>
                      <c r="D2">
                        <v>-1</v>
                      </c>
                      <c r="E2">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="3" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A3">
                        <v>-2</v>
                      </c>
                      <c r="B3">
                        <v>2</v>
                      </c>
                      <c r="C3">
                        <v>3</v>
                      </c>
                      <c r="D3">
                        <v>-1</v>
                      </c>
                      <c r="E3">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="4" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A4">
                        <v>-2</v>
                      </c>
                      <c r="B4">
                        <v>2</v>
                      </c>
                      <c r="C4">
                        <v>3</v>
                      </c>
                      <c r="D4">
                        <v>-1</v>
                      </c>
                      <c r="E4">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="5" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A5">
                        <v>-2</v>
                      </c>
                      <c r="B5">
                        <v>2</v>
                      </c>
                      <c r="C5">
                        <v>3</v>
                      </c>
                      <c r="D5">
                        <v>-1</v>
                      </c>
                      <c r="E5">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="6" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A6">
                        <v>-2</v>
                      </c>
                      <c r="B6">
                        <v>2</v>
                      </c>
                      <c r="C6">
                        <v>3</v>
                      </c>
                      <c r="D6">
                        <v>-1</v>
                      </c>
                      <c r="E6">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="7" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A7">
                        <v>-2</v>
                      </c>
                      <c r="B7">
                        <v>2</v>
                      </c>
                      <c r="C7">
                        <v>3</v>
                      </c>
                      <c r="D7">
                        <v>-1</v>
                      </c>
                      <c r="E7">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="8" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A8">
                        <v>-2</v>
                      </c>
                      <c r="B8">
                        <v>2</v>
                      </c>
                      <c r="C8">
                        <v>3</v>
                      </c>
                      <c r="D8">
                        <v>-1</v>
                      </c>
                      <c r="E8">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="9" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A9">
                        <v>-2</v>
                      </c>
                      <c r="B9">
                        <v>2</v>
                      </c>
                      <c r="C9">
                        <v>3</v>
                      </c>
                      <c r="D9">
                        <v>-1</v>
                      </c>
                      <c r="E9">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="10" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A10">
                        <v>-2</v>
                      </c>
                      <c r="B10">
                        <v>2</v>
                      </c>
                      <c r="C10">
                        <v>3</v>
                      </c>
                      <c r="D10">
                        <v>-1</v>
                      </c>
                      <c r="E10">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="11" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A11">
                        <v>-2</v>
                      </c>
                      <c r="B11">
                        <v>2</v>
                      </c>
                      <c r="C11">
                        <v>3</v>
                      </c>
                      <c r="D11">
                        <v>-1</v>
                      </c>
                      <c r="E11">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="12" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A12">
                        <v>-2</v>
                      </c>
                      <c r="B12">
                        <v>2</v>
                      </c>
                      <c r="C12">
                        <v>3</v>
                      </c>
                      <c r="D12">
                        <v>-1</v>
                      </c>
                      <c r="E12">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="13" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A13">
                        <v>-2</v>
                      </c>
                      <c r="B13">
                        <v>2</v>
                      </c>
                      <c r="C13">
                        <v>3</v>
                      </c>
                      <c r="D13">
                        <v>-1</v>
                      </c>
                      <c r="E13">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="14" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A14">
                        <v>-2</v>
                      </c>
                      <c r="B14">
                        <v>2</v>
                      </c>
                      <c r="C14">
                        <v>3</v>
                      </c>
                      <c r="D14">
                        <v>-1</v>
                      </c>
                      <c r="E14">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="15" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A15">
                        <v>-2</v>
                      </c>
                      <c r="B15">
                        <v>2</v>
                      </c>
                      <c r="C15">
                        <v>3</v>
                      </c>
                      <c r="D15">
                        <v>-1</v>
                      </c>
                      <c r="E15">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="16" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A16">
                        <v>-2</v>
                      </c>
                      <c r="B16">
                        <v>2</v>
                      </c>
                      <c r="C16">
                        <v>3</v>
                      </c>
                      <c r="D16">
                        <v>-1</v>
                      </c>
                      <c r="E16">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="17" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A17">
                        <v>-2</v>
                      </c>
                      <c r="B17">
                        <v>2</v>
                      </c>
                      <c r="C17">
                        <v>3</v>
                      </c>
                      <c r="D17">
                        <v>-1</v>
                      </c>
                      <c r="E17">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="18" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A18">
                        <v>-2</v>
                      </c>
                      <c r="B18">
                        <v>2</v>
                      </c>
                      <c r="C18">
                        <v>3</v>
                      </c>
                      <c r="D18">
                        <v>-1</v>
                      </c>
                      <c r="E18">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="19" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A19">
                        <v>-2</v>
                      </c>
                      <c r="B19">
                        <v>2</v>
                      </c>
                      <c r="C19">
                        <v>3</v>
                      </c>
                      <c r="D19">
                        <v>-1</v>
                      </c>
                      <c r="E19">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="20" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A20">
                        <v>-2</v>
                      </c>
                      <c r="B20">
                        <v>2</v>
                      </c>
                      <c r="C20">
                        <v>3</v>
                      </c>
                      <c r="D20">
                        <v>-1</v>
                      </c>
                      <c r="E20">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="21" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A21">
                        <v>-2</v>
                      </c>
                      <c r="B21">
                        <v>2</v>
                      </c>
                      <c r="C21">
                        <v>3</v>
                      </c>
                      <c r="D21">
                        <v>-1</v>
                      </c>
                      <c r="E21">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="22" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A22">
                        <v>-2</v>
                      </c>
                      <c r="B22">
                        <v>2</v>
                      </c>
                      <c r="C22">
                        <v>3</v>
                      </c>
                      <c r="D22">
                        <v>-1</v>
                      </c>
                      <c r="E22">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="23" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A23">
                        <v>-2</v>
                      </c>
                      <c r="B23">
                        <v>2</v>
                      </c>
                      <c r="C23">
                        <v>3</v>
                      </c>
                      <c r="D23">
                        <v>-1</v>
                      </c>
                      <c r="E23">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="24" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A24">
                        <v>-2</v>
                      </c>
                      <c r="B24">
                        <v>2</v>
                      </c>
                      <c r="C24">
                        <v>3</v>
                      </c>
                      <c r="D24">
                        <v>-1</v>
                      </c>
                      <c r="E24">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="25" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A25">
                        <v>-2</v>
                      </c>
                      <c r="B25">
                        <v>2</v>
                      </c>
                      <c r="C25">
                        <v>3</v>
                      </c>
                      <c r="D25">
                        <v>-1</v>
                      </c>
                      <c r="E25">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="26" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A26">
                        <v>-2</v>
                      </c>
                      <c r="B26">
                        <v>2</v>
                      </c>
                      <c r="C26">
                        <v>3</v>
                      </c>
                      <c r="D26">
                        <v>-1</v>
                      </c>
                      <c r="E26">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="27" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A27">
                        <v>-2</v>
                      </c>
                      <c r="B27">
                        <v>2</v>
                      </c>
                      <c r="C27">
                        <v>3</v>
                      </c>
                      <c r="D27">
                        <v>-1</v>
                      </c>
                      <c r="E27">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="28" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A28">
                        <v>-2</v>
                      </c>
                      <c r="B28">
                        <v>2</v>
                      </c>
                      <c r="C28">
                        <v>3</v>
                      </c>
                      <c r="D28">
                        <v>-1</v>
                      </c>
                      <c r="E28">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="29" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A29">
                        <v>-2</v>
                      </c>
                      <c r="B29">
                        <v>2</v>
                      </c>
                      <c r="C29">
                        <v>3</v>
                      </c>
                      <c r="D29">
                        <v>-1</v>
                      </c>
                      <c r="E29">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="30" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A30">
                        <v>-2</v>
                      </c>
                      <c r="B30">
                        <v>2</v>
                      </c>
                      <c r="C30">
                        <v>3</v>
                      </c>
                      <c r="D30">
                        <v>-1</v>
                      </c>
                      <c r="E30">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="31" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A31">
                        <v>-2</v>
                      </c>
                      <c r="B31">
                        <v>2</v>
                      </c>
                      <c r="C31">
                        <v>3</v>
                      </c>
                      <c r="D31">
                        <v>-1</v>
                      </c>
                      <c r="E31">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="32" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A32">
                        <v>-2</v>
                      </c>
                      <c r="B32">
                        <v>2</v>
                      </c>
                      <c r="C32">
                        <v>3</v>
                      </c>
                      <c r="D32">
                        <v>-1</v>
                      </c>
                      <c r="E32">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="33" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A33">
                        <v>-2</v>
                      </c>
                      <c r="B33">
                        <v>2</v>
                      </c>
                      <c r="C33">
                        <v>3</v>
                      </c>
                      <c r="D33">
                        <v>-1</v>
                      </c>
                      <c r="E33">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="34" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A34">
                        <v>-2</v>
                      </c>
                      <c r="B34">
                        <v>2</v>
                      </c>
                      <c r="C34">
                        <v>3</v>
                      </c>
                      <c r="D34">
                        <v>-1</v>
                      </c>
                      <c r="E34">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="35" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A35">
                        <v>-2</v>
                      </c>
                      <c r="B35">
                        <v>2</v>
                      </c>
                      <c r="C35">
                        <v>3</v>
                      </c>
                      <c r="D35">
                        <v>-1</v>
                      </c>
                      <c r="E35">
                        <v>0</v>
                      </c>
                    </row>
                    <row r="36" spans="1:5" x14ac:dyDescent="0.25">
                      <c r="A36">
                        <v>-2</v>
                      </c>
                      <c r="B36">
                        <v>2</v>
                      </c>
                      <c r="C36">
                        <v>3</v>
                      </c>
                      <c r="D36">
                        <v>-1</v>
                      </c>
                      <c r="E36">
                        <v>0</v>
                      </c>
                    </row>
                  </sheetData>
                  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
                  <extLst>
                    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
                      <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="1"/>
                          <x14:colorNegative theme="9"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="8"/>
                          <x14:colorFirst theme="4"/>
                          <x14:colorLast theme="5"/>
                          <x14:colorHigh theme="6"/>
                          <x14:colorLow theme="7"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A36:E36</xm:f>
                              <xm:sqref>F36</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="3"/>
                          <x14:colorNegative theme="9"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="8"/>
                          <x14:colorFirst theme="4"/>
                          <x14:colorLast theme="5"/>
                          <x14:colorHigh theme="6"/>
                          <x14:colorLow theme="7"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A35:E35</xm:f>
                              <xm:sqref>F35</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries rgb="FF00B050"/>
                          <x14:colorNegative rgb="FFFF0000"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers rgb="FF0070C0"/>
                          <x14:colorFirst rgb="FFFFC000"/>
                          <x14:colorLast rgb="FFFFC000"/>
                          <x14:colorHigh rgb="FF00B050"/>
                          <x14:colorLow rgb="FFFF0000"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A34:E34</xm:f>
                              <xm:sqref>F34</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries rgb="FFC6EFCE"/>
                          <x14:colorNegative rgb="FFFFC7CE"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers rgb="FF8CADD6"/>
                          <x14:colorFirst rgb="FFFFDC47"/>
                          <x14:colorLast rgb="FFFFEB9C"/>
                          <x14:colorHigh rgb="FF60D276"/>
                          <x14:colorLow rgb="FFFF5367"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A33:E33</xm:f>
                              <xm:sqref>F33</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries rgb="FF5687C2"/>
                          <x14:colorNegative rgb="FFFFB620"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers rgb="FFD70077"/>
                          <x14:colorFirst rgb="FF777777"/>
                          <x14:colorLast rgb="FF359CEB"/>
                          <x14:colorHigh rgb="FF56BE79"/>
                          <x14:colorLow rgb="FFFF5055"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A32:E32</xm:f>
                              <xm:sqref>F32</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries rgb="FF5F5F5F"/>
                          <x14:colorNegative rgb="FFFFB620"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers rgb="FFD70077"/>
                          <x14:colorFirst rgb="FF5687C2"/>
                          <x14:colorLast rgb="FF359CEB"/>
                          <x14:colorHigh rgb="FF56BE79"/>
                          <x14:colorLow rgb="FFFF5055"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A31:E31</xm:f>
                              <xm:sqref>F31</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries rgb="FF0070C0"/>
                          <x14:colorNegative rgb="FF000000"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers rgb="FF000000"/>
                          <x14:colorFirst rgb="FF000000"/>
                          <x14:colorLast rgb="FF000000"/>
                          <x14:colorHigh rgb="FF000000"/>
                          <x14:colorLow rgb="FF000000"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A30:E30</xm:f>
                              <xm:sqref>F30</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries rgb="FF376092"/>
                          <x14:colorNegative rgb="FFD00000"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers rgb="FFD00000"/>
                          <x14:colorFirst rgb="FFD00000"/>
                          <x14:colorLast rgb="FFD00000"/>
                          <x14:colorHigh rgb="FFD00000"/>
                          <x14:colorLow rgb="FFD00000"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A29:E29</xm:f>
                              <xm:sqref>F29</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries rgb="FF000000"/>
                          <x14:colorNegative rgb="FF0070C0"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers rgb="FF0070C0"/>
                          <x14:colorFirst rgb="FF0070C0"/>
                          <x14:colorLast rgb="FF0070C0"/>
                          <x14:colorHigh rgb="FF0070C0"/>
                          <x14:colorLow rgb="FF0070C0"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A28:E28</xm:f>
                              <xm:sqref>F28</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries rgb="FF323232"/>
                          <x14:colorNegative rgb="FFD00000"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers rgb="FFD00000"/>
                          <x14:colorFirst rgb="FFD00000"/>
                          <x14:colorLast rgb="FFD00000"/>
                          <x14:colorHigh rgb="FFD00000"/>
                          <x14:colorLow rgb="FFD00000"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A27:E27</xm:f>
                              <xm:sqref>F27</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="1" tint="0.34998626667073579"/>
                          <x14:colorNegative theme="0" tint="-0.249977111117893"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="0" tint="-0.249977111117893"/>
                          <x14:colorFirst theme="0" tint="-0.249977111117893"/>
                          <x14:colorLast theme="0" tint="-0.249977111117893"/>
                          <x14:colorHigh theme="0" tint="-0.249977111117893"/>
                          <x14:colorLow theme="0" tint="-0.249977111117893"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A26:E26</xm:f>
                              <xm:sqref>F26</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="1" tint="0.499984740745262"/>
                          <x14:colorNegative theme="1" tint="0.249977111117893"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="1" tint="0.249977111117893"/>
                          <x14:colorFirst theme="1" tint="0.249977111117893"/>
                          <x14:colorLast theme="1" tint="0.249977111117893"/>
                          <x14:colorHigh theme="1" tint="0.249977111117893"/>
                          <x14:colorLow theme="1" tint="0.249977111117893"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A25:E25</xm:f>
                              <xm:sqref>F25</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="9" tint="0.39997558519241921"/>
                          <x14:colorNegative theme="0" tint="-0.499984740745262"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="9" tint="0.79998168889431442"/>
                          <x14:colorFirst theme="9" tint="-0.249977111117893"/>
                          <x14:colorLast theme="9" tint="-0.249977111117893"/>
                          <x14:colorHigh theme="9" tint="-0.499984740745262"/>
                          <x14:colorLow theme="9" tint="-0.499984740745262"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A24:E24</xm:f>
                              <xm:sqref>F24</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="8" tint="0.39997558519241921"/>
                          <x14:colorNegative theme="0" tint="-0.499984740745262"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="8" tint="0.79998168889431442"/>
                          <x14:colorFirst theme="8" tint="-0.249977111117893"/>
                          <x14:colorLast theme="8" tint="-0.249977111117893"/>
                          <x14:colorHigh theme="8" tint="-0.499984740745262"/>
                          <x14:colorLow theme="8" tint="-0.499984740745262"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A23:E23</xm:f>
                              <xm:sqref>F23</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="7" tint="0.39997558519241921"/>
                          <x14:colorNegative theme="0" tint="-0.499984740745262"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="7" tint="0.79998168889431442"/>
                          <x14:colorFirst theme="7" tint="-0.249977111117893"/>
                          <x14:colorLast theme="7" tint="-0.249977111117893"/>
                          <x14:colorHigh theme="7" tint="-0.499984740745262"/>
                          <x14:colorLow theme="7" tint="-0.499984740745262"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A22:E22</xm:f>
                              <xm:sqref>F22</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="6" tint="0.39997558519241921"/>
                          <x14:colorNegative theme="0" tint="-0.499984740745262"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="6" tint="0.79998168889431442"/>
                          <x14:colorFirst theme="6" tint="-0.249977111117893"/>
                          <x14:colorLast theme="6" tint="-0.249977111117893"/>
                          <x14:colorHigh theme="6" tint="-0.499984740745262"/>
                          <x14:colorLow theme="6" tint="-0.499984740745262"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A21:E21</xm:f>
                              <xm:sqref>F21</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="5" tint="0.39997558519241921"/>
                          <x14:colorNegative theme="0" tint="-0.499984740745262"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="5" tint="0.79998168889431442"/>
                          <x14:colorFirst theme="5" tint="-0.249977111117893"/>
                          <x14:colorLast theme="5" tint="-0.249977111117893"/>
                          <x14:colorHigh theme="5" tint="-0.499984740745262"/>
                          <x14:colorLow theme="5" tint="-0.499984740745262"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A20:E20</xm:f>
                              <xm:sqref>F20</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="4" tint="0.39997558519241921"/>
                          <x14:colorNegative theme="0" tint="-0.499984740745262"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="4" tint="0.79998168889431442"/>
                          <x14:colorFirst theme="4" tint="-0.249977111117893"/>
                          <x14:colorLast theme="4" tint="-0.249977111117893"/>
                          <x14:colorHigh theme="4" tint="-0.499984740745262"/>
                          <x14:colorLow theme="4" tint="-0.499984740745262"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A19:E19</xm:f>
                              <xm:sqref>F19</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="9"/>
                          <x14:colorNegative theme="4"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="9" tint="-0.249977111117893"/>
                          <x14:colorFirst theme="9" tint="-0.249977111117893"/>
                          <x14:colorLast theme="9" tint="-0.249977111117893"/>
                          <x14:colorHigh theme="9" tint="-0.249977111117893"/>
                          <x14:colorLow theme="9" tint="-0.249977111117893"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A18:E18</xm:f>
                              <xm:sqref>F18</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="8"/>
                          <x14:colorNegative theme="9"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="8" tint="-0.249977111117893"/>
                          <x14:colorFirst theme="8" tint="-0.249977111117893"/>
                          <x14:colorLast theme="8" tint="-0.249977111117893"/>
                          <x14:colorHigh theme="8" tint="-0.249977111117893"/>
                          <x14:colorLow theme="8" tint="-0.249977111117893"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A17:E17</xm:f>
                              <xm:sqref>F17</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="7"/>
                          <x14:colorNegative theme="8"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="7" tint="-0.249977111117893"/>
                          <x14:colorFirst theme="7" tint="-0.249977111117893"/>
                          <x14:colorLast theme="7" tint="-0.249977111117893"/>
                          <x14:colorHigh theme="7" tint="-0.249977111117893"/>
                          <x14:colorLow theme="7" tint="-0.249977111117893"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A16:E16</xm:f>
                              <xm:sqref>F16</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="6"/>
                          <x14:colorNegative theme="7"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="6" tint="-0.249977111117893"/>
                          <x14:colorFirst theme="6" tint="-0.249977111117893"/>
                          <x14:colorLast theme="6" tint="-0.249977111117893"/>
                          <x14:colorHigh theme="6" tint="-0.249977111117893"/>
                          <x14:colorLow theme="6" tint="-0.249977111117893"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A15:E15</xm:f>
                              <xm:sqref>F15</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="5"/>
                          <x14:colorNegative theme="6"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="5" tint="-0.249977111117893"/>
                          <x14:colorFirst theme="5" tint="-0.249977111117893"/>
                          <x14:colorLast theme="5" tint="-0.249977111117893"/>
                          <x14:colorHigh theme="5" tint="-0.249977111117893"/>
                          <x14:colorLow theme="5" tint="-0.249977111117893"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A14:E14</xm:f>
                              <xm:sqref>F14</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="4"/>
                          <x14:colorNegative theme="5"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="4" tint="-0.249977111117893"/>
                          <x14:colorFirst theme="4" tint="-0.249977111117893"/>
                          <x14:colorLast theme="4" tint="-0.249977111117893"/>
                          <x14:colorHigh theme="4" tint="-0.249977111117893"/>
                          <x14:colorLow theme="4" tint="-0.249977111117893"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A13:E13</xm:f>
                              <xm:sqref>F13</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="9" tint="-0.249977111117893"/>
                          <x14:colorNegative theme="4"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="4" tint="-0.249977111117893"/>
                          <x14:colorFirst theme="4" tint="-0.249977111117893"/>
                          <x14:colorLast theme="4" tint="-0.249977111117893"/>
                          <x14:colorHigh theme="4" tint="-0.249977111117893"/>
                          <x14:colorLow theme="4" tint="-0.249977111117893"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A12:E12</xm:f>
                              <xm:sqref>F12</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="8" tint="-0.249977111117893"/>
                          <x14:colorNegative theme="9"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="9" tint="-0.249977111117893"/>
                          <x14:colorFirst theme="9" tint="-0.249977111117893"/>
                          <x14:colorLast theme="9" tint="-0.249977111117893"/>
                          <x14:colorHigh theme="9" tint="-0.249977111117893"/>
                          <x14:colorLow theme="9" tint="-0.249977111117893"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A11:E11</xm:f>
                              <xm:sqref>F11</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="7" tint="-0.249977111117893"/>
                          <x14:colorNegative theme="8"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="8" tint="-0.249977111117893"/>
                          <x14:colorFirst theme="8" tint="-0.249977111117893"/>
                          <x14:colorLast theme="8" tint="-0.249977111117893"/>
                          <x14:colorHigh theme="8" tint="-0.249977111117893"/>
                          <x14:colorLow theme="8" tint="-0.249977111117893"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A10:E10</xm:f>
                              <xm:sqref>F10</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="6" tint="-0.249977111117893"/>
                          <x14:colorNegative theme="7"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="7" tint="-0.249977111117893"/>
                          <x14:colorFirst theme="7" tint="-0.249977111117893"/>
                          <x14:colorLast theme="7" tint="-0.249977111117893"/>
                          <x14:colorHigh theme="7" tint="-0.249977111117893"/>
                          <x14:colorLow theme="7" tint="-0.249977111117893"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A9:E9</xm:f>
                              <xm:sqref>F9</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="5" tint="-0.249977111117893"/>
                          <x14:colorNegative theme="6"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="6" tint="-0.249977111117893"/>
                          <x14:colorFirst theme="6" tint="-0.249977111117893"/>
                          <x14:colorLast theme="6" tint="-0.249977111117893"/>
                          <x14:colorHigh theme="6" tint="-0.249977111117893"/>
                          <x14:colorLow theme="6" tint="-0.249977111117893"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A8:E8</xm:f>
                              <xm:sqref>F8</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="4" tint="-0.249977111117893"/>
                          <x14:colorNegative theme="5"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="5" tint="-0.249977111117893"/>
                          <x14:colorFirst theme="5" tint="-0.249977111117893"/>
                          <x14:colorLast theme="5" tint="-0.249977111117893"/>
                          <x14:colorHigh theme="5" tint="-0.249977111117893"/>
                          <x14:colorLow theme="5" tint="-0.249977111117893"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A7:E7</xm:f>
                              <xm:sqref>F7</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="9" tint="-0.499984740745262"/>
                          <x14:colorNegative theme="4"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="9" tint="-0.499984740745262"/>
                          <x14:colorFirst theme="9" tint="0.39997558519241921"/>
                          <x14:colorLast theme="9" tint="0.39997558519241921"/>
                          <x14:colorHigh theme="9"/>
                          <x14:colorLow theme="9"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A6:E6</xm:f>
                              <xm:sqref>F6</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="8" tint="-0.499984740745262"/>
                          <x14:colorNegative theme="9"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="8" tint="-0.499984740745262"/>
                          <x14:colorFirst theme="8" tint="0.39997558519241921"/>
                          <x14:colorLast theme="8" tint="0.39997558519241921"/>
                          <x14:colorHigh theme="8"/>
                          <x14:colorLow theme="8"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A5:E5</xm:f>
                              <xm:sqref>F5</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="7" tint="-0.499984740745262"/>
                          <x14:colorNegative theme="8"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="7" tint="-0.499984740745262"/>
                          <x14:colorFirst theme="7" tint="0.39997558519241921"/>
                          <x14:colorLast theme="7" tint="0.39997558519241921"/>
                          <x14:colorHigh theme="7"/>
                          <x14:colorLow theme="7"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A4:E4</xm:f>
                              <xm:sqref>F4</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="6" tint="-0.499984740745262"/>
                          <x14:colorNegative theme="7"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="6" tint="-0.499984740745262"/>
                          <x14:colorFirst theme="6" tint="0.39997558519241921"/>
                          <x14:colorLast theme="6" tint="0.39997558519241921"/>
                          <x14:colorHigh theme="6"/>
                          <x14:colorLow theme="6"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A3:E3</xm:f>
                              <xm:sqref>F3</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="5" tint="-0.499984740745262"/>
                          <x14:colorNegative theme="6"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="5" tint="-0.499984740745262"/>
                          <x14:colorFirst theme="5" tint="0.39997558519241921"/>
                          <x14:colorLast theme="5" tint="0.39997558519241921"/>
                          <x14:colorHigh theme="5"/>
                          <x14:colorLow theme="5"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A2:E2</xm:f>
                              <xm:sqref>F2</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                        <x14:sparklineGroup displayEmptyCellsAs="gap">
                          <x14:colorSeries theme="4" tint="-0.499984740745262"/>
                          <x14:colorNegative theme="5"/>
                          <x14:colorAxis rgb="FF000000"/>
                          <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
                          <x14:colorFirst theme="4" tint="0.39997558519241921"/>
                          <x14:colorLast theme="4" tint="0.39997558519241921"/>
                          <x14:colorHigh theme="4"/>
                          <x14:colorLow theme="4"/>
                          <x14:sparklines>
                            <x14:sparkline>
                              <xm:f>Sheet1!A1:E1</xm:f>
                              <xm:sqref>F1</xm:sqref>
                            </x14:sparkline>
                          </x14:sparklines>
                        </x14:sparklineGroup>
                      </x14:sparklineGroups>
                    </ext>
                  </extLst>
                </worksheet>
                """)

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)
