###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2025, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO

from xlsxwriter.exceptions import ThemeFileError
from xlsxwriter.workbook import Workbook


class TestThemeExceptions(unittest.TestCase):
    """
    Test workbook exceptions when calling use_custom_theme().

    """

    def test_theme_exception01(self):
        """Test use_custom_theme() exceptions"""
        workbook = Workbook()

        theme = StringIO("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""")

        # No exception.
        workbook.use_custom_theme(theme)

    def test_theme_exception02(self):
        """Test use_custom_theme() exceptions"""
        workbook = Workbook()

        theme = StringIO("<invalid></invalid>")

        with self.assertRaises(ThemeFileError):
            workbook.use_custom_theme(theme)

    def test_theme_exception03(self):
        """Test use_custom_theme() exceptions"""
        workbook = Workbook()

        theme = StringIO("<a:blipFill>")

        with self.assertRaises(ThemeFileError):
            workbook.use_custom_theme(theme)
