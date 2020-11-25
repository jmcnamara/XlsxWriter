###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2020, John McNamara, jmcnamara@cpan.org
#

import os
import tempfile
import unittest
import warnings
from ...workbook import Workbook
from ...exceptions import FileCreateError


class TestCloseWithException(unittest.TestCase):
    """
    Test the Workbook close() exception.

    """
    def test_non_existent_dir(self):
        """Test the _check_sheetname() method"""

        self.workbook = Workbook('non_existent_path/test.xlsx')
        self.workbook.add_worksheet()

        with self.assertRaises(FileCreateError):
            self.workbook.close()

    def test_workbook_closes_all_handles(self):
        """Test that close() closes all file handles"""

        filepath = tempfile.mktemp()

        with warnings.catch_warnings(record=True) as warnings_emitted:
            workbook = Workbook(filepath, dict(constant_memory=True))
            workbook.close()
            del workbook

        os.unlink(filepath)
        self.assertFalse(warnings_emitted)
