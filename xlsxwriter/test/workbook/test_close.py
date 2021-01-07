###############################################################################
#
# Tests for XlsxWriter.
#
# Copyright (c), 2013-2021, John McNamara, jmcnamara@cpan.org
#

import unittest
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
