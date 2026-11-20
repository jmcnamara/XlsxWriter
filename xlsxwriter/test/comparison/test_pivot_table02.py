###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

from xlsxwriter.workbook import Workbook

from ..excel_comparison_test import ExcelComparisonTest


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename("pivot_table02.xlsx")
        self.set_text_file("autofilter_data.txt")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file with a pivot table with a row field and a value field."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        # Open a text file with the pivot table source data.
        with open(self.txt_filename, encoding="utf-8") as textfile:
            # Read the text file and write it to the worksheet.
            for row, line in enumerate(textfile):
                # Split the input data based on whitespace.
                data = line.strip("\n").split()

                # Convert the number data from the text file.
                for i, item in enumerate(data):
                    try:
                        data[i] = float(item)
                    except ValueError:
                        pass

                # Write out the row data.
                worksheet.write_row(row, 0, data)

        worksheet = workbook.add_worksheet()
        worksheet.add_pivot_table(
            2,
            0,
            {
                "source": "Sheet1!A1:D51",
                "rows": ["Region"],
                "values": [{"field": "Volume"}],
            },
        )

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_with_constant_memory(self):
        """Test the creation of a simple XlsxWriter file with a pivot table with field names for a constant_memory data source."""

        # Ignore the parts that differ due to constant_memory inline strings.
        self.ignore_files = [
            "[Content_Types].xml",
            "xl/_rels/workbook.xml.rels",
            "xl/sharedStrings.xml",
            "xl/worksheets/sheet1.xml",
        ]

        workbook = Workbook(
            self.got_filename, {"constant_memory": True, "in_memory": False}
        )
        worksheet = workbook.add_worksheet()

        # Open a text file with the pivot table source data.
        with open(self.txt_filename, encoding="utf-8") as textfile:
            # Read the text file and write it to the worksheet.
            for row, line in enumerate(textfile):
                # Split the input data based on whitespace.
                data = line.strip("\n").split()

                # Convert the number data from the text file.
                for i, item in enumerate(data):
                    try:
                        data[i] = float(item)
                    except ValueError:
                        pass

                # Write out the row data.
                worksheet.write_row(row, 0, data)

        worksheet = workbook.add_worksheet()
        worksheet.add_pivot_table(
            2,
            0,
            {
                "source": "Sheet1!A1:D51",
                "field_names": ["Region", "Item", "Volume", "Month"],
                "rows": ["Region"],
                "values": [{"field": "Volume"}],
            },
        )

        workbook.close()

        self.assertExcelEqual()
