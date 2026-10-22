###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
###############################################################################

import unittest
from io import BytesIO
from xml.etree import ElementTree
from zipfile import ZipFile

from xlsxwriter import Workbook


class TestConditionalFormatTextQuotes(unittest.TestCase):
    """Test string literals in generated conditional format formulas."""

    def test_text_quotes(self):
        """Test each text criterion with quotes and XML special characters."""
        cases = [
            ('"', '""', 1),
            ('say "hello"', 'say ""hello""', 11),
            ('a""b & <c>', 'a""""b & <c>', 10),
        ]
        for value, escaped, length in cases:
            expected = [
                ("containing", "containsText", f'NOT(ISERROR(SEARCH("{escaped}",A1)))'),
                (
                    "not containing",
                    "notContainsText",
                    f'ISERROR(SEARCH("{escaped}",A1))',
                ),
                ("begins with", "beginsWith", f'LEFT(A1,{length})="{escaped}"'),
                ("ends with", "endsWith", f'RIGHT(A1,{length})="{escaped}"'),
            ]
            with self.subTest(value=value):
                output = BytesIO()
                with Workbook(output) as workbook:
                    worksheet = workbook.add_worksheet()
                    worksheet.write("A1", value)
                    cell_format = workbook.add_format({"bg_color": "#FFFF00"})
                    for criterion, _, _ in expected:
                        result = worksheet.conditional_format(
                            "A1:A4",
                            {
                                "type": "text",
                                "criteria": criterion,
                                "value": value,
                                "format": cell_format,
                            },
                        )
                        self.assertEqual(result, 0)

                with ZipFile(output) as archive:
                    xml = ElementTree.fromstring(
                        archive.read("xl/worksheets/sheet1.xml")
                    )
                rules = xml.findall(".//{*}cfRule")
                self.assertEqual(len(rules), len(expected))
                for rule, (criterion, rule_type, formula) in zip(rules, expected):
                    with self.subTest(criterion=criterion):
                        self.assertEqual(rule.attrib["type"], rule_type)
                        self.assertEqual(rule.attrib["text"], value)
                        self.assertEqual(rule.find("{*}formula").text, formula)

    def test_explicit_formulas_unchanged(self):
        """Test that cell/formula criteria preserve user-supplied formulas."""
        output = BytesIO()
        with Workbook(output) as workbook:
            worksheet = workbook.add_worksheet()
            worksheet.conditional_format(
                "A1",
                {"type": "cell", "criteria": "==", "value": '"say ""hello"""'},
            )
            worksheet.conditional_format(
                "A1",
                {"type": "formula", "criteria": '=A1="say ""hello"""'},
            )
        with ZipFile(output) as archive:
            xml = ElementTree.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        formulas = [
            rule.find("{*}formula").text for rule in xml.findall(".//{*}cfRule")
        ]
        self.assertEqual(formulas, ['"say ""hello"""', 'A1="say ""hello"""'])
