###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO
from ...workbook import Workbook
from ...exceptions import DuplicateTableName


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
