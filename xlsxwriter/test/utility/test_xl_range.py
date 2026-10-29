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

from xlsxwriter.utility import xl_range, xl_range_abs


class TestUtility(unittest.TestCase):
    """
    Test xl_range() utility function.

    """

    def test_xl_range(self):
        """Test xl_range()"""

        self.assertEqual("A1:A10", xl_range(0, 0, 9, 0))
        self.assertEqual("C2:C9", xl_range(1, 2, 8, 2))
        self.assertEqual("A1:E4", xl_range(0, 0, 3, 4))
        self.assertEqual("A1", xl_range(0, 0, 0, 0))
        self.assertEqual("A1:B1", xl_range(0, 0, 0, 1))
        self.assertEqual("C1:J1", xl_range(0, 2, 0, 9))
        self.assertEqual("A2:A3", xl_range(1, 0, 2, 0))
        self.assertEqual("A10:Y2", xl_range(9, 0, 1, 24))
        self.assertEqual("Z8:AA10", xl_range(7, 25, 9, 26))
        self.assertEqual("IU2:IV2", xl_range(1, 254, 1, 255))
        self.assertEqual("IW2:XFD1", xl_range(1, 256, 0, 16383))
        self.assertEqual("A1:XFE1048577", xl_range(0, 0, 1048576, 16384))

        # Ignore the warnings for negative values.
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            self.assertEqual("", xl_range(-1, 0, 0, 0))
            self.assertEqual("", xl_range(0, -1, 0, 0))
            self.assertEqual("", xl_range(0, 0, -1, 0))
            self.assertEqual("", xl_range(0, 0, 0, -1))

    def test_xl_range_abs(self):
        """Test xl_range_abs()"""

        self.assertEqual("$A$1:$A$10", xl_range_abs(0, 0, 9, 0))
        self.assertEqual("$C$2:$C$9", xl_range_abs(1, 2, 8, 2))
        self.assertEqual("$A$1:$E$4", xl_range_abs(0, 0, 3, 4))
        self.assertEqual("$A$1", xl_range_abs(0, 0, 0, 0))
        self.assertEqual("$A$1:$B$1", xl_range_abs(0, 0, 0, 1))
        self.assertEqual("$C$1:$J$1", xl_range_abs(0, 2, 0, 9))
        self.assertEqual("$A$2:$A$3", xl_range_abs(1, 0, 2, 0))
        self.assertEqual("$A$10:$Y$2", xl_range_abs(9, 0, 1, 24))
        self.assertEqual("$Z$8:$AA$10", xl_range_abs(7, 25, 9, 26))
        self.assertEqual("$IU$2:$IV$2", xl_range_abs(1, 254, 1, 255))
        self.assertEqual("$IW$2:$XFD$1", xl_range_abs(1, 256, 0, 16383))
        self.assertEqual("$A$1:$XFE$1048577", xl_range_abs(0, 0, 1048576, 16384))

        # Ignore the warnings for negative values.
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            self.assertEqual("", xl_range_abs(-1, 0, 0, 0))
            self.assertEqual("", xl_range_abs(0, -1, 0, 0))
            self.assertEqual("", xl_range_abs(0, 0, -1, 0))
            self.assertEqual("", xl_range_abs(0, 0, 0, -1))
