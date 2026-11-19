###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import unittest
from datetime import date, datetime, timedelta, timezone

from xlsxwriter.workbook import Workbook


class TestSetCustomProperty(unittest.TestCase):
    """
    Test the Workbook set_custom_property() method with dates.

    """

    def setUp(self):
        self.workbook = Workbook()

    def tearDown(self):
        self.workbook.fileclosed = 1

    def test_set_custom_property_date(self):
        """Test set_custom_property() with a date object."""

        self.workbook.set_custom_property("Checked", date(2018, 1, 1))

        exp = ("Checked", "2018-01-01T00:00:00Z", "date")
        got = self.workbook.custom_properties[0]

        self.assertEqual(exp, got)

    def test_set_custom_property_timezone(self):
        """Test set_custom_property() with a timezone aware datetime."""

        tz = timezone(timedelta(hours=5, minutes=30))
        self.workbook.set_custom_property(
            "Checked", datetime(2018, 1, 1, 0, 15, tzinfo=tz)
        )

        exp = ("Checked", "2017-12-31T18:45:00Z", "date")
        got = self.workbook.custom_properties[0]

        self.assertEqual(exp, got)
