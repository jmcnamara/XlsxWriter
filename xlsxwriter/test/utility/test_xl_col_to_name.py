###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import unittest
import warnings

from xlsxwriter.utility import xl_col_to_name


class TestUtility(unittest.TestCase):
    """
    Test xl_col_to_name() utility function.

    """

    def test_xl_col_to_name(self):
        """Test xl_col_to_name()"""

        self.assertEqual("A", xl_col_to_name(0))
        self.assertEqual("B", xl_col_to_name(1))
        self.assertEqual("C", xl_col_to_name(2))
        self.assertEqual("J", xl_col_to_name(9))
        self.assertEqual("Y", xl_col_to_name(24))
        self.assertEqual("Z", xl_col_to_name(25))
        self.assertEqual("AA", xl_col_to_name(26))
        self.assertEqual("IU", xl_col_to_name(254))
        self.assertEqual("IV", xl_col_to_name(255))
        self.assertEqual("IW", xl_col_to_name(256))
        self.assertEqual("XFD", xl_col_to_name(16383))
        self.assertEqual("XFE", xl_col_to_name(16384))

        # Ignore the warnings for negative values.
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            self.assertEqual("", xl_col_to_name(-1))

    def test_xl_col_to_name_abs(self):
        """Test xl_col_to_name() with absolute references"""

        self.assertEqual("$A", xl_col_to_name(0, True))

        # Ignore the warnings for negative values.
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            self.assertEqual("", xl_col_to_name(-1, True))
