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

from xlsxwriter.exceptions import OverlappingRange, PivotTableError
from xlsxwriter.workbook import Workbook


class TestPivotTableErrors(unittest.TestCase):
    """
    Test the pivot table option and data source validation.

    """

    def setUp(self):
        self.workbook = Workbook(BytesIO())
        self.data_worksheet = self.workbook.add_worksheet("Data")
        self.worksheet = self.workbook.add_worksheet()
        self.options = {"source": ["Data", 0, 0, 50, 3], "rows": ["Region"]}

    def tearDown(self):
        self.workbook.fileclosed = True

    def test_options(self):
        """Test invalid pivot table options."""

        # Check for an unknown option.
        with self.assertRaises(PivotTableError):
            self.worksheet.add_pivot_table(2, 0, {**self.options, "colour": "red"})

        # Check for an unknown value field option.
        values = [{"field": "Volume", "colour": "red"}]
        with self.assertRaises(PivotTableError):
            self.worksheet.add_pivot_table(2, 0, {**self.options, "values": values})

        # Check for an unknown value field function.
        values = [{"field": "Volume", "function": "median"}]
        with self.assertRaises(PivotTableError):
            self.worksheet.add_pivot_table(2, 0, {**self.options, "values": values})

        # Check for a value field without a field name.
        values = [{"function": "count"}]
        with self.assertRaises(PivotTableError):
            self.worksheet.add_pivot_table(2, 0, {**self.options, "values": values})

        # Check for a field list that isn't a list of strings.
        with self.assertRaises(PivotTableError):
            self.worksheet.add_pivot_table(2, 0, {**self.options, "rows": "Region"})

        # Check for an unknown style and layout.
        with self.assertRaises(PivotTableError):
            self.worksheet.add_pivot_table(
                2, 0, {**self.options, "style": "Pivot Style Light 29"}
            )

        with self.assertRaises(PivotTableError):
            self.worksheet.add_pivot_table(2, 0, {**self.options, "layout": "grid"})

    def test_repeated_fields(self):
        """Test fields used more than once in the row, column or filter fields."""

        # Check that a field can't be used twice in the same axis.
        with self.assertRaises(PivotTableError):
            self.worksheet.add_pivot_table(
                2, 0, {**self.options, "rows": ["Region", "Region"]}
            )

        # Check that a field can't be used in more than one of the row, column
        # and filter axes.
        with self.assertRaises(PivotTableError):
            self.worksheet.add_pivot_table(
                2, 0, {**self.options, "columns": ["Region"]}
            )

        with self.assertRaises(PivotTableError):
            self.worksheet.add_pivot_table(
                2, 0, {**self.options, "columns": ["Item"], "filters": ["Item"]}
            )

        # Check that a field can be used as a row field and a value field, and
        # in several value fields.
        values = [{"field": "Region"}, {"field": "Region"}]
        self.worksheet.add_pivot_table(2, 0, {**self.options, "values": values})

    def test_data_source(self):
        """Test invalid pivot table data sources."""

        # Check for a missing data source.
        with self.assertRaises(PivotTableError):
            self.worksheet.add_pivot_table(2, 0, {"rows": ["Region"]})

        # Check for an invalid data source range string or list.
        for source in ["Data!A1", "Data!A1:XFE9", ["Data", 0, 0, 50], 1]:
            with self.assertRaises(PivotTableError):
                self.worksheet.add_pivot_table(2, 0, {"source": source})

        # Check for a data source with a header row but no data rows.
        with self.assertRaises(PivotTableError):
            self.worksheet.add_pivot_table(2, 0, {"source": ["Data", 0, 0, 0, 3]})

        # Check for user defined field names that don't match the number of
        # source columns.
        with self.assertRaises(PivotTableError):
            self.worksheet.add_pivot_table(
                2,
                0,
                {**self.options, "field_names": ["Region", "Item", "Volume"]},
            )

        # Check for a data source with the rows or columns reversed.
        for source in [
            ["Data", 50, 0, 0, 3],
            ["Data", 0, 3, 50, 0],
            "Data!D51:A1",
        ]:
            with self.assertRaises(PivotTableError):
                self.worksheet.add_pivot_table(2, 0, {"source": source})

        # Check for a data source outside Excel's row or column limits.
        for source in [
            ["Data", 0, 0, 1048576, 3],
            ["Data", 0, 0, 50, 16384],
            ["Data", -1, 0, 50, 3],
        ]:
            with self.assertRaises(PivotTableError):
                self.worksheet.add_pivot_table(2, 0, {"source": source})

    def test_position(self):
        """Test pivot table positions outside the worksheet limits."""

        # Check for a position outside Excel's row or column limits.
        for row, col in [(1048576, 0), (0, 16384), (-1, 0)]:
            with self.assertRaises(PivotTableError):
                self.worksheet.add_pivot_table(row, col, self.options)

        # Check for filter fields that move the pivot table body past the last
        # row. The body is below the filters and a blank row.
        options = {**self.options, "filters": ["Month"]}
        with self.assertRaises(PivotTableError):
            self.worksheet.add_pivot_table(1048574, 0, options)

        # Check that the body can start in the last row.
        self.worksheet.add_pivot_table(1048573, 0, options)

    def test_chartsheet(self):
        """Test a pivot table added to a chartsheet."""

        # Check for a pivot table added to a chartsheet.
        chartsheet = self.workbook.add_chartsheet()

        with self.assertRaises(PivotTableError):
            chartsheet.add_pivot_table(0, 0, self.options)

    def test_overlap(self):
        """Test pivot tables that overlap other worksheet objects."""

        # Check for a pivot table on a worksheet table.
        worksheet = self.workbook.add_worksheet()
        worksheet.add_table(2, 0, 8, 3)

        with self.assertRaises(PivotTableError):
            worksheet.add_pivot_table(2, 0, self.options)

        # Check for a worksheet table on a pivot table.
        worksheet = self.workbook.add_worksheet()
        worksheet.add_pivot_table(2, 0, self.options)

        with self.assertRaises(OverlappingRange):
            worksheet.add_table(2, 0, 8, 3)

        # Check for a pivot table on a merged range.
        worksheet = self.workbook.add_worksheet()
        worksheet.merge_range(2, 0, 2, 3, "Merged")

        with self.assertRaises(PivotTableError):
            worksheet.add_pivot_table(2, 0, self.options)

        # Check for a merged range on a pivot table.
        worksheet = self.workbook.add_worksheet()
        worksheet.add_pivot_table(2, 0, self.options)

        with self.assertRaises(OverlappingRange):
            worksheet.merge_range(2, 0, 2, 3, "Merged")

        # Check for two pivot tables at the same position.
        worksheet = self.workbook.add_worksheet()
        worksheet.add_pivot_table(2, 0, self.options)

        with self.assertRaises(PivotTableError):
            worksheet.add_pivot_table(2, 0, self.options)

        # Check for a pivot table whose filter fields overlap a previous pivot
        # table. The body is below the filters and a blank row.
        worksheet = self.workbook.add_worksheet()
        worksheet.add_pivot_table(4, 0, self.options)

        with self.assertRaises(PivotTableError):
            worksheet.add_pivot_table(2, 0, {**self.options, "filters": ["Month"]})

        # Check for a pivot table on its own source data.
        with self.assertRaises(PivotTableError):
            self.data_worksheet.add_pivot_table(10, 1, self.options)

        # Check that pivot tables, tables and merged ranges apart from each
        # other are allowed.
        worksheet = self.workbook.add_worksheet()
        worksheet.add_table(0, 6, 4, 7)
        worksheet.merge_range(0, 1, 0, 3, "Merged")
        worksheet.add_pivot_table(2, 0, self.options)
        worksheet.add_pivot_table(2, 3, self.options)
        self.data_worksheet.add_pivot_table(10, 4, self.options)


class TestPivotTableSaveErrors(unittest.TestCase):
    """
    Test the pivot table validation when the workbook is saved.

    """

    def save_pivot_table(self, headers, options, workbook_options=None):
        # Create a workbook with a pivot table data source on the first
        # worksheet and a pivot table on the second, then save it.
        workbook = Workbook(BytesIO(), workbook_options)

        worksheet = workbook.add_worksheet()
        worksheet.write_row(0, 0, headers)
        worksheet.write(1, 0, "East")

        worksheet = workbook.add_worksheet()
        worksheet.add_pivot_table(2, 0, options)

        try:
            workbook.close()
        finally:
            workbook.fileclosed = True

    def test_source_sheet(self):
        """Test a data source on an unknown worksheet or a chartsheet."""

        # Check for a data source on a worksheet that doesn't exist.
        with self.assertRaises(PivotTableError):
            self.save_pivot_table(
                ["Region"], {"source": ["Sheet3", 0, 0, 1, 0], "rows": ["Region"]}
            )

        # Check for a data source on a chartsheet.
        workbook = Workbook(BytesIO())
        chart = workbook.add_chart({"type": "column"})
        chart.add_series({"values": "=Sheet1!$A$2:$A$3"})

        workbook.add_worksheet()
        chartsheet = workbook.add_chartsheet()
        chartsheet.set_chart(chart)

        worksheet = workbook.add_worksheet()
        worksheet.add_pivot_table(
            2, 0, {"source": ["Chart1", 0, 0, 1, 0], "rows": ["Region"]}
        )

        with self.assertRaises(PivotTableError):
            workbook.close()

        workbook.fileclosed = True

    def test_constant_memory_source(self):
        """Test a data source on a constant_memory worksheet."""

        workbook_options = {"constant_memory": True, "in_memory": False}

        # Check for a data source on a constant_memory worksheet without user
        # defined field names.
        with self.assertRaises(PivotTableError):
            self.save_pivot_table(
                ["Region"],
                {"source": ["Sheet1", 0, 0, 1, 0], "rows": ["Region"]},
                workbook_options,
            )

        # Check that user defined field names are allowed.
        self.save_pivot_table(
            ["Region"],
            {
                "source": ["Sheet1", 0, 0, 1, 0],
                "rows": ["Region"],
                "field_names": ["Region"],
            },
            workbook_options,
        )

    def test_user_field_names(self):
        """Test user defined field names on a standard worksheet."""

        # Check that user defined field names must match the header row on a
        # standard worksheet.
        options = {"source": ["Sheet1", 0, 0, 1, 1], "rows": ["Region"]}

        self.save_pivot_table(
            ["Region", "Volume"], {**options, "field_names": ["Region", "Volume"]}
        )

        with self.assertRaises(PivotTableError):
            self.save_pivot_table(
                ["Region", "Volume"], {**options, "field_names": ["Region", "Item"]}
            )

    def test_field_names(self):
        """Test field names that don't match the source data headers."""

        # Check for a field name that doesn't match a source header in the
        # row, column, filter or value fields. Field names are case sensitive.
        for options in [
            {"rows": ["Legion"]},
            {"columns": ["Legion"]},
            {"filters": ["Legion"]},
            {"values": [{"field": "Legion"}]},
            {"rows": ["REGION"]},
        ]:
            with self.assertRaises(PivotTableError):
                self.save_pivot_table(
                    ["Region", "Volume"],
                    {"source": ["Sheet1", 0, 0, 1, 1], **options},
                )

    def test_value_field_names(self):
        """Test duplicate value field names."""

        workbook = Workbook(BytesIO())
        worksheet = workbook.add_worksheet()
        options = {"source": ["Data", 0, 0, 50, 3]}

        # Check for two value fields with the same custom name, or with custom
        # names that differ only by case.
        for names in [("Total", "Total"), ("Total", "TOTAL")]:
            values = [{"field": "Volume", "name": name} for name in names]

            with self.assertRaises(PivotTableError):
                worksheet.add_pivot_table(2, 0, {**options, "values": values})

        # Check for a custom name that is the same as a default name.
        values = [{"field": "Volume"}, {"field": "Volume", "name": "Sum of Volume"}]

        with self.assertRaises(PivotTableError):
            worksheet.add_pivot_table(2, 0, {**options, "values": values})

        # Check that repeated default names are allowed since they are made
        # unique.
        values = [{"field": "Volume"}, {"field": "Volume"}]
        worksheet.add_pivot_table(2, 0, {**options, "values": values})

        workbook.fileclosed = True

        # Check for a custom name that is the same as a source header, ignoring
        # case.
        with self.assertRaises(PivotTableError):
            self.save_pivot_table(
                ["Region", "Volume"],
                {
                    "source": ["Sheet1", 0, 0, 1, 1],
                    "rows": ["Region"],
                    "values": [{"field": "Volume", "name": "VOLUME"}],
                },
            )

    def test_headers(self):
        """Test invalid header cells in the data source."""

        # Check for a blank header cell, a missing header cell, a missing
        # header row or a duplicate header, ignoring case.
        for headers in [
            ["Region", "", "Volume"],
            ["Region", "Volume"],
            [],
            ["Region", "Item", "Region"],
            ["Region", "Item", "REGION"],
        ]:
            with self.assertRaises(PivotTableError):
                self.save_pivot_table(
                    headers, {"source": ["Sheet1", 0, 0, 1, 2], "rows": ["Region"]}
                )

        # Check for number, boolean and formula header cells, which aren't
        # supported.
        for headers in [
            ["Region", 2024, "Volume"],
            ["Region", True, "Volume"],
            ["Region", "=1+1", "Volume"],
        ]:
            with self.assertRaises(PivotTableError):
                self.save_pivot_table(
                    headers, {"source": ["Sheet1", 0, 0, 1, 2], "rows": ["Region"]}
                )

        # Check that string and rich string header cells are allowed.
        workbook = Workbook(BytesIO())
        bold = workbook.add_format({"bold": True})

        worksheet = workbook.add_worksheet()
        worksheet.write_row(0, 0, ["Region", "Item"])
        worksheet.write_rich_string(0, 2, bold, "Vol", "ume")
        worksheet.write(1, 0, "East")

        worksheet = workbook.add_worksheet()
        worksheet.add_pivot_table(
            2,
            0,
            {
                "source": ["Sheet1", 0, 0, 1, 2],
                "rows": ["Region"],
                "values": [{"field": "Volume"}],
            },
        )

        workbook.close()


if __name__ == "__main__":
    unittest.main()
