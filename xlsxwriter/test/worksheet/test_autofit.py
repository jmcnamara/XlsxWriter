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


class TestAutofit(unittest.TestCase):
    """Test autofit column properties in a completed XLSX file."""

    def setUp(self):
        self.output = BytesIO()
        self.workbook = Workbook(self.output)
        self.worksheet = self.workbook.add_worksheet()

    def get_column(self):
        """Close the workbook and read its serialized column attributes."""
        self.workbook.close()
        with ZipFile(self.output) as archive:
            root = ElementTree.fromstring(archive.read("xl/worksheets/sheet1.xml"))

        namespace = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
        columns = root.findall(f"{namespace}cols/{namespace}col")
        self.assertEqual(len(columns), 1)
        return columns[0].attrib

    def test_autofit_after_set_column(self):
        """Widening a user-set column must not hide it (issue #1219)."""
        self.worksheet.set_column(0, 0, 2)
        self.worksheet.write_string(0, 0, "hello world")
        self.worksheet.autofit()

        self.assertEqual(
            self.get_column(),
            {
                "min": "1",
                "max": "1",
                "width": "11.140625",
                "bestFit": "1",
                "customWidth": "1",
            },
        )

    def test_autofit_after_set_column_pixels(self):
        """A user-set pixel width can also be widened without hiding it."""
        self.worksheet.set_column_pixels(0, 0, 19)
        self.worksheet.write_string(0, 0, "hello world")
        self.worksheet.autofit()

        self.assertEqual(
            self.get_column(),
            {
                "min": "1",
                "max": "1",
                "width": "11.140625",
                "bestFit": "1",
                "customWidth": "1",
            },
        )

    def test_autofit_preserves_user_minimum(self):
        """A wider user-set column retains its width and flags."""
        self.worksheet.set_column(0, 0, 20)
        self.worksheet.write_string(0, 0, "hello world")
        self.worksheet.autofit()

        self.assertEqual(
            self.get_column(),
            {"min": "1", "max": "1", "width": "20.7109375", "customWidth": "1"},
        )

    def test_autofit_format_only_column(self):
        """Autofit preserves a column's format and outline properties."""
        bold = self.workbook.add_format({"bold": True})
        self.worksheet.set_column(0, 0, None, bold, {"level": 2, "collapsed": True})
        self.worksheet.write_string(0, 0, "hello world")
        self.worksheet.autofit()

        self.assertEqual(
            self.get_column(),
            {
                "min": "1",
                "max": "1",
                "width": "11.140625",
                "style": "1",
                "bestFit": "1",
                "customWidth": "1",
                "outlineLevel": "2",
                "collapsed": "1",
            },
        )

    def test_autofit_preserves_hidden_column(self):
        """An explicitly hidden column stays hidden when autofitted."""
        self.worksheet.set_column(0, 0, 2, None, {"hidden": True})
        self.worksheet.write_string(0, 0, "hello world")
        self.worksheet.autofit()

        self.assertEqual(
            self.get_column(),
            {
                "min": "1",
                "max": "1",
                "width": "11.140625",
                "hidden": "1",
                "bestFit": "1",
                "customWidth": "1",
            },
        )

    def test_autofit_preserves_hidden_column_without_width(self):
        """A hidden column without an explicit width also stays hidden."""
        self.worksheet.set_column(0, 0, None, None, {"hidden": True})
        self.worksheet.write_string(0, 0, "hello world")
        self.worksheet.autofit()

        self.assertEqual(
            self.get_column(),
            {
                "min": "1",
                "max": "1",
                "width": "11.140625",
                "hidden": "1",
                "bestFit": "1",
                "customWidth": "1",
            },
        )

    def test_autofit_preserves_hidden_user_minimum(self):
        """A hidden column also retains a wider user-set minimum width."""
        self.worksheet.set_column(0, 0, 20, None, {"hidden": True})
        self.worksheet.write_string(0, 0, "hello world")
        self.worksheet.autofit()

        self.assertEqual(
            self.get_column(),
            {
                "min": "1",
                "max": "1",
                "width": "20.7109375",
                "hidden": "1",
                "customWidth": "1",
            },
        )

    def test_autofit_repeated(self):
        """Repeated autofit calls must not hide an autofitted column."""
        self.worksheet.write_string(0, 0, "hello world")
        self.worksheet.autofit()
        self.worksheet.autofit()

        self.assertEqual(
            self.get_column(),
            {
                "min": "1",
                "max": "1",
                "width": "11.140625",
                "bestFit": "1",
                "customWidth": "1",
            },
        )

    def test_autofit_repeated_widens_column(self):
        """Further autofit calls can widen a column after more data is added."""
        self.worksheet.set_column(0, 0, 2)
        self.worksheet.write_string(0, 0, "hello")
        self.worksheet.autofit()
        self.worksheet.write_string(1, 0, "hello world")
        self.worksheet.autofit()
        self.worksheet.autofit()

        self.assertEqual(
            self.get_column(),
            {
                "min": "1",
                "max": "1",
                "width": "11.140625",
                "bestFit": "1",
                "customWidth": "1",
            },
        )

    def test_autofit_max_width(self):
        """An existing column is widened only to the autofit maximum."""
        self.worksheet.set_column(0, 0, 2)
        self.worksheet.write_string(0, 0, "W" * 100)
        self.worksheet.autofit(100)

        self.assertEqual(
            self.get_column(),
            {
                "min": "1",
                "max": "1",
                "width": "14.28515625",
                "bestFit": "1",
                "customWidth": "1",
            },
        )

    def test_autofit_repeated_with_smaller_max_width(self):
        """A subsequent autofit can apply a smaller maximum width."""
        self.worksheet.write_string(0, 0, "W" * 100)
        self.worksheet.autofit()
        self.worksheet.autofit(100)

        self.assertEqual(
            self.get_column(),
            {
                "min": "1",
                "max": "1",
                "width": "14.28515625",
                "bestFit": "1",
                "customWidth": "1",
            },
        )

    def test_autofit_max_width_preserves_user_minimum(self):
        """The autofit maximum does not reduce a wider user-set width."""
        self.worksheet.set_column(0, 0, 20)
        self.worksheet.write_string(0, 0, "W" * 100)
        self.worksheet.autofit(100)

        self.assertEqual(
            self.get_column(),
            {"min": "1", "max": "1", "width": "20.7109375", "customWidth": "1"},
        )
