###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import BytesIO, StringIO
from xml.etree import ElementTree
from zipfile import ZipFile

from xlsxwriter.exceptions import DuplicateTableName, OverlappingRange
from xlsxwriter.workbook import Workbook


class TestAddTable(unittest.TestCase):
    """
    Test exceptions with add_table().

    """

    def test_duplicate_table_name(self):
        """Test that adding 2 tables with the same name raises an exception."""

        fh = StringIO()
        workbook = Workbook()
        workbook._set_filehandle(fh)

        worksheet = workbook.add_worksheet()

        worksheet.add_table("B1:F3", {"name": "SalesData"})
        worksheet.add_table("B4:F7", {"name": "SalesData"})

        self.assertRaises(DuplicateTableName, workbook._prepare_tables)

        workbook.fileclosed = True

    def test_rejected_table_range_can_be_reused(self):
        """Test that invalid tables don't prevent a corrected table being added."""
        cases = [
            ("A1:B2", {"unknown": True}),
            ("A1:B1", {}),
            ("A1:B2", {"name": "Sales Data"}),
            ("A1:B2", {"name": "1Sales"}),
            ("A1:B2", {"name": "Sales" * 52}),
            ("A1:B2", {"name": "A1"}),
            ("A1:B2", {"name": "R1C1"}),
            ("A1:B2", {"columns": [{"header": "Sales"}, {"header": "sales"}]}),
            ("A1:B2", {"columns": [{"header": "Column2"}]}),
        ]

        for cell_range, options in cases:
            with self.subTest(cell_range=cell_range, options=options):
                fh = BytesIO()
                with Workbook(fh, {"in_memory": True}) as workbook:
                    worksheet = workbook.add_worksheet()
                    with self.assertWarns(UserWarning):
                        self.assertEqual(worksheet.add_table(cell_range, options), -2)

                    self.assertEqual(worksheet.add_table("A1:B2"), 0)

                with ZipFile(fh) as xlsx:
                    table = ElementTree.fromstring(xlsx.read("xl/tables/table1.xml"))
                    self.assertEqual(table.get("ref"), "A1:B2")
                    self.assertNotIn("xl/tables/table2.xml", xlsx.namelist())

    def test_rejected_table_range_can_be_merged(self):
        """Test that a rejected table doesn't prevent merging its range."""
        fh = BytesIO()
        with Workbook(fh, {"in_memory": True}) as workbook:
            worksheet = workbook.add_worksheet()
            with self.assertWarns(UserWarning):
                self.assertEqual(worksheet.add_table("A1:B2", {"unknown": True}), -2)

            self.assertEqual(worksheet.merge_range("A1:B2", "Merged"), 0)

        with ZipFile(fh) as xlsx:
            self.assertNotIn("xl/tables/table1.xml", xlsx.namelist())
            sheet = ElementTree.fromstring(xlsx.read("xl/worksheets/sheet1.xml"))
            merge = sheet.find("{*}mergeCells/{*}mergeCell")
            self.assertEqual(merge.get("ref"), "A1:B2")

    def test_overlapping_table_does_not_reserve_free_cells(self):
        """Test that a rejected overlap doesn't reserve its non-overlapping cells."""
        for existing_range in ("table", "merge"):
            with self.subTest(existing_range=existing_range):
                fh = BytesIO()
                with Workbook(fh, {"in_memory": True}) as workbook:
                    worksheet = workbook.add_worksheet()
                    if existing_range == "table":
                        self.assertEqual(worksheet.add_table("B1:C2"), 0)
                    else:
                        self.assertEqual(worksheet.merge_range("B1:C2", "Merged"), 0)

                    with self.assertRaises(OverlappingRange):
                        worksheet.add_table("A1:B2")

                    self.assertEqual(worksheet.add_table("A1:A2"), 0)

                with ZipFile(fh) as xlsx:
                    ranges = [
                        ElementTree.fromstring(xlsx.read(name)).get("ref")
                        for name in xlsx.namelist()
                        if name.startswith("xl/tables/")
                    ]
                    expected = ["B1:C2", "A1:A2"]
                    if existing_range == "merge":
                        expected = ["A1:A2"]
                    self.assertEqual(ranges, expected)
