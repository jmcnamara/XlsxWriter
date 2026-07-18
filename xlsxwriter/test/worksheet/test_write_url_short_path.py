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

from xlsxwriter.url import Url
from xlsxwriter.workbook import Workbook


class TestWriteUrlShortPath(unittest.TestCase):
    """Short file/external paths must not IndexError in _is_relative_path."""

    def test_is_relative_path_short_strings(self):
        self.assertTrue(Url._is_relative_path(""))
        self.assertTrue(Url._is_relative_path("a"))
        self.assertTrue(Url._is_relative_path("ab"))
        self.assertFalse(Url._is_relative_path("C:"))
        self.assertFalse(Url._is_relative_path("C:\\foo"))
        self.assertFalse(Url._is_relative_path(r"\\server\share"))

    def test_write_url_file_short_path(self):
        workbook = Workbook(BytesIO(), {"in_memory": True})
        workbook.default_url_format = None
        worksheet = workbook.add_worksheet()
        worksheet.write_url("A1", "file:///a")
        workbook.close()
        self.assertEqual(worksheet.hyperlinks[0][0]._link, "a")

    def test_write_url_external_short_path(self):
        workbook = Workbook(BytesIO(), {"in_memory": True})
        workbook.default_url_format = None
        worksheet = workbook.add_worksheet()
        worksheet.write_url("A1", "external:a")
        workbook.close()
        self.assertEqual(worksheet.hyperlinks[0][0]._link, "a")
