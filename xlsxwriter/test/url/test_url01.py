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

from xlsxwriter.url import Url
from xlsxwriter.workbook import Workbook


class TestUrl(unittest.TestCase):
    """Test cases for the Url class."""

    def test_tip_length(self):
        """Test valid tooltip lengths through both URL entry points."""
        for tip in ("", "x" * 255):
            with self.subTest(length=len(tip)):
                url = Url("https://example.com")
                url.tip = tip
                self.assertEqual(url.tip, tip)

                url = Url.from_options({"url": "https://example.com", "tip": tip})
                self.assertEqual(url.tip, tip)

    def test_tip_too_long(self):
        """Reject an overlong tooltip without changing the existing value."""
        url = Url("https://example.com")
        url.tip = "Existing tooltip"

        with self.assertRaisesRegex(ValueError, "Hyperlink tool tip exceeds"):
            url.tip = "x" * 256

        self.assertEqual(url.tip, "Existing tooltip")

    def test_tip_from_options_too_long(self):
        """Reject an overlong tooltip supplied in an options dictionary."""
        with self.assertRaisesRegex(ValueError, "Hyperlink tool tip exceeds"):
            Url.from_options({"url": "https://example.com", "tip": "x" * 256})

    def test_write_url_tip_length(self):
        """Write a tooltip at the Excel limit for external and internal links."""
        for link in ("https://example.com", "internal:Sheet1!B1"):
            with self.subTest(link=link):
                output = BytesIO()
                tip = "x" * 255
                with Workbook(output) as workbook:
                    worksheet = workbook.add_worksheet()
                    self.assertEqual(worksheet.write_url("A1", link, tip=tip), 0)

                with ZipFile(output) as archive:
                    sheet = ElementTree.fromstring(
                        archive.read("xl/worksheets/sheet1.xml")
                    )
                    hyperlink = sheet.find("{*}hyperlinks/{*}hyperlink")
                    self.assertEqual(hyperlink.attrib["tooltip"], tip)

    def test_write_url_tip_too_long(self):
        """Reject an overlong tooltip before registering the hyperlink."""
        with Workbook(BytesIO()) as workbook:
            worksheet = workbook.add_worksheet()
            with self.assertRaisesRegex(ValueError, "Hyperlink tool tip exceeds"):
                worksheet.write_url("A1", "https://example.com", tip="x" * 256)

            self.assertEqual(worksheet.hlink_count, 0)
            self.assertEqual(worksheet.hyperlinks, {})
