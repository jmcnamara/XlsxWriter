###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import BytesIO
from xml.etree import ElementTree
from zipfile import ZipFile

from xlsxwriter.workbook import Workbook


class TestRichString(unittest.TestCase):
    def test_literal_tags_in_inline_string(self):
        output = BytesIO()
        with Workbook(output, {"constant_memory": True}) as workbook:
            worksheet = workbook.add_worksheet()
            worksheet.write_string("A1", "<r>abc</r>")

        with ZipFile(output) as xlsx:
            sheet = ElementTree.fromstring(xlsx.read("xl/worksheets/sheet1.xml"))

        self.assertEqual(sheet.findtext(".//{*}is/{*}t"), "<r>abc</r>")

    def test_literal_and_formatted_shared_strings(self):
        # The literal text is identical to the XML for the formatted cell.
        rich_xml = (
            '<r><t>a</t></r><r><rPr><b/><sz val="11"/><color theme="1"/>'
            '<rFont val="Calibri"/><family val="2"/><scheme val="minor"/>'
            "</rPr><t>b</t></r>"
        )

        output = BytesIO()
        with Workbook(output) as workbook:
            worksheet = workbook.add_worksheet()
            bold = workbook.add_format({"bold": True})
            for row in (0, 1):
                worksheet.write_string(row, 0, rich_xml)
                worksheet.write_rich_string(row, 1, "a", bold, "b")

        with ZipFile(output) as xlsx:
            shared_xml = xlsx.read("xl/sharedStrings.xml").decode()
            sheet = ElementTree.fromstring(xlsx.read("xl/worksheets/sheet1.xml"))

        indices = [cell.text for cell in sheet.findall(".//{*}c/{*}v")]
        self.assertEqual(indices, ["0", "1", "0", "1"])
        strings = ElementTree.fromstring(shared_xml)
        self.assertEqual(strings[0].findtext("{*}t"), rich_xml)
        self.assertIn(f"<si>{rich_xml}</si>", shared_xml)
