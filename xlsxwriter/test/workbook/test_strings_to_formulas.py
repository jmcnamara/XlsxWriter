###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import BytesIO
from xml.etree import ElementTree
from zipfile import ZipFile

from xlsxwriter.workbook import Workbook


class TestStringsToFormulas(unittest.TestCase):
    """Check formula conversion in the completed workbook."""

    def test_array_formula_conversion(self):
        for constant_memory in (False, True):
            for enabled in (None, True, False):
                with self.subTest(constant_memory=constant_memory, enabled=enabled):
                    options = {"constant_memory": constant_memory}
                    if enabled is not None:
                        options["strings_to_formulas"] = enabled
                    output = BytesIO()
                    with Workbook(output, options) as workbook:
                        worksheet = workbook.add_worksheet()
                        self.assertEqual(worksheet.write("A1", "=1+2"), 0)
                        self.assertEqual(worksheet.write("B1", "{=1+2}"), 0)
                        self.assertEqual(
                            worksheet.write_row("A2", ["=1+2", "{=1+2}"]), 0
                        )
                        self.assertEqual(worksheet.write_column("A3", ["{=1+2}"]), 0)

                    ns = {
                        "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    }
                    with ZipFile(output) as archive:
                        root = ElementTree.fromstring(
                            archive.read("xl/worksheets/sheet1.xml")
                        )
                        cells = root.findall(".//m:c", ns)
                        self.assertEqual(len(cells), 5)
                        if enabled is not False:
                            for cell in cells:
                                self.assertEqual(cell.find("m:f", ns).text, "1+2")
                            for cell in (cells[1], cells[3], cells[4]):
                                self.assertEqual(cell.find("m:f", ns).get("t"), "array")
                        else:
                            strings = []
                            if not constant_memory:
                                shared = ElementTree.fromstring(
                                    archive.read("xl/sharedStrings.xml")
                                )
                                strings = [
                                    node.text for node in shared.findall("m:si/m:t", ns)
                                ]
                            values = []
                            for cell in cells:
                                self.assertIsNone(cell.find("m:f", ns))
                                if constant_memory:
                                    values.append(cell.find("m:is/m:t", ns).text)
                                else:
                                    values.append(
                                        strings[int(cell.find("m:v", ns).text)]
                                    )
                            self.assertEqual(
                                values, ["=1+2", "{=1+2}", "=1+2", "{=1+2}", "{=1+2}"]
                            )

    def test_explicit_formulas_ignore_conversion_option(self):
        output = BytesIO()
        with Workbook(output, {"strings_to_formulas": False}) as workbook:
            worksheet = workbook.add_worksheet()
            self.assertEqual(worksheet.write_formula("A1", "=1+2"), 0)
            self.assertEqual(worksheet.write_formula("A2", "{=1+2}"), 0)
            self.assertEqual(worksheet.write_array_formula("A3", "{=1+2}"), 0)
            self.assertEqual(worksheet.write_string("A4", "{=1+2}"), 0)
        ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        with ZipFile(output) as archive:
            root = ElementTree.fromstring(archive.read("xl/worksheets/sheet1.xml"))
            self.assertEqual(len(root.findall(".//m:f", ns)), 3)
            self.assertEqual(
                [node.text for node in root.findall(".//m:f", ns)], ["1+2"] * 3
            )
            self.assertIsNone(root.find(".//m:c[@r='A4']/m:f", ns))
