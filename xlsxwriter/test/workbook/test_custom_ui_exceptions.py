###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import unittest

from xlsxwriter.exceptions import CustomUIFileError
from xlsxwriter.workbook import Workbook


class TestCustomUIExceptions(unittest.TestCase):
    """
    Test workbook exceptions when calling add_custom_ui().

    """

    def test_custom_ui_exception01(self):
        """Second file with the same format."""
        workbook = Workbook()
        custom_ui = "xlsxwriter/test/comparison/xlsx_files/customUI14-01.xml"

        workbook.add_custom_ui(custom_ui)

        with self.assertRaises(CustomUIFileError):
            workbook.add_custom_ui(custom_ui)
