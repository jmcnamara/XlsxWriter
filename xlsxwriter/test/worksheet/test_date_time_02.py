###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2023, John McNamara, jmcnamara@cpan.org
#

import datetime
import unittest

from ...worksheet import Worksheet


class TestConvertDateTime(unittest.TestCase):
    """
    Test the Worksheet _convert_date_time() method against dates extracted
    from Excel.

    """

    def setUp(self):
        self.worksheet = Worksheet()

        # Dates and corresponding numbers from an Excel file.
        self.excel_dates = [
            ("1899-12-31T", 0),
            # 1900-1-1 fails for datetime.datetime due to a difference in the
            # way it handles time only values and the way Excel does.
            # ('1900-01-01T', 1),
            ("1900-02-27T", 58),
            ("1900-02-28T", 59),
            ("1900-03-01T", 61),
            ("1900-03-02T", 62),
            ("1900-03-11T", 71),
            ("1900-04-08T", 99),
            ("1900-09-12T", 256),
            ("1901-05-03T", 489),
            ("1901-10-13T", 652),
            ("1902-02-15T", 777),
            ("1902-06-06T", 888),
            ("1902-09-25T", 999),
            ("1902-09-27T", 1001),
            ("1903-04-26T", 1212),
            ("1903-08-05T", 1313),
            ("1903-12-31T", 1461),
            ("1904-01-01T", 1462),
            ("1904-02-28T", 1520),
            ("1904-02-29T", 1521),
            ("1904-03-01T", 1522),
            ("1907-02-27T", 2615),
            ("1907-02-28T", 2616),
            ("1907-03-01T", 2617),
            ("1907-03-02T", 2618),
            ("1907-03-03T", 2619),
            ("1907-03-04T", 2620),
            ("1907-03-05T", 2621),
            ("1907-03-06T", 2622),
            ("1999-01-01T", 36161),
            ("1999-01-31T", 36191),
            ("1999-02-01T", 36192),
            ("1999-02-28T", 36219),
            ("1999-03-01T", 36220),
            ("1999-03-31T", 36250),
            ("1999-04-01T", 36251),
            ("1999-04-30T", 36280),
            ("1999-05-01T", 36281),
            ("1999-05-31T", 36311),
            ("1999-06-01T", 36312),
            ("1999-06-30T", 36341),
            ("1999-07-01T", 36342),
            ("1999-07-31T", 36372),
            ("1999-08-01T", 36373),
            ("1999-08-31T", 36403),
            ("1999-09-01T", 36404),
            ("1999-09-30T", 36433),
            ("1999-10-01T", 36434),
            ("1999-10-31T", 36464),
            ("1999-11-01T", 36465),
            ("1999-11-30T", 36494),
            ("1999-12-01T", 36495),
            ("1999-12-31T", 36525),
            ("2000-01-01T", 36526),
            ("2000-01-31T", 36556),
            ("2000-02-01T", 36557),
            ("2000-02-29T", 36585),
            ("2000-03-01T", 36586),
            ("2000-03-31T", 36616),
            ("2000-04-01T", 36617),
            ("2000-04-30T", 36646),
            ("2000-05-01T", 36647),
            ("2000-05-31T", 36677),
            ("2000-06-01T", 36678),
            ("2000-06-30T", 36707),
            ("2000-07-01T", 36708),
            ("2000-07-31T", 36738),
            ("2000-08-01T", 36739),
            ("2000-08-31T", 36769),
            ("2000-09-01T", 36770),
            ("2000-09-30T", 36799),
            ("2000-10-01T", 36800),
            ("2000-10-31T", 36830),
            ("2000-11-01T", 36831),
            ("2000-11-30T", 36860),
            ("2000-12-01T", 36861),
            ("2000-12-31T", 36891),
            ("2001-01-01T", 36892),
            ("2001-01-31T", 36922),
            ("2001-02-01T", 36923),
            ("2001-02-28T", 36950),
            ("2001-03-01T", 36951),
            ("2001-03-31T", 36981),
            ("2001-04-01T", 36982),
            ("2001-04-30T", 37011),
            ("2001-05-01T", 37012),
            ("2001-05-31T", 37042),
            ("2001-06-01T", 37043),
            ("2001-06-30T", 37072),
            ("2001-07-01T", 37073),
            ("2001-07-31T", 37103),
            ("2001-08-01T", 37104),
            ("2001-08-31T", 37134),
            ("2001-09-01T", 37135),
            ("2001-09-30T", 37164),
            ("2001-10-01T", 37165),
            ("2001-10-31T", 37195),
            ("2001-11-01T", 37196),
            ("2001-11-30T", 37225),
            ("2001-12-01T", 37226),
            ("2001-12-31T", 37256),
            ("2400-01-01T", 182623),
            ("2400-01-31T", 182653),
            ("2400-02-01T", 182654),
            ("2400-02-29T", 182682),
            ("2400-03-01T", 182683),
            ("2400-03-31T", 182713),
            ("2400-04-01T", 182714),
            ("2400-04-30T", 182743),
            ("2400-05-01T", 182744),
            ("2400-05-31T", 182774),
            ("2400-06-01T", 182775),
            ("2400-06-30T", 182804),
            ("2400-07-01T", 182805),
            ("2400-07-31T", 182835),
            ("2400-08-01T", 182836),
            ("2400-08-31T", 182866),
            ("2400-09-01T", 182867),
            ("2400-09-30T", 182896),
            ("2400-10-01T", 182897),
            ("2400-10-31T", 182927),
            ("2400-11-01T", 182928),
            ("2400-11-30T", 182957),
            ("2400-12-01T", 182958),
            ("2400-12-31T", 182988),
            ("4000-01-01T", 767011),
            ("4000-01-31T", 767041),
            ("4000-02-01T", 767042),
            ("4000-02-29T", 767070),
            ("4000-03-01T", 767071),
            ("4000-03-31T", 767101),
            ("4000-04-01T", 767102),
            ("4000-04-30T", 767131),
            ("4000-05-01T", 767132),
            ("4000-05-31T", 767162),
            ("4000-06-01T", 767163),
            ("4000-06-30T", 767192),
            ("4000-07-01T", 767193),
            ("4000-07-31T", 767223),
            ("4000-08-01T", 767224),
            ("4000-08-31T", 767254),
            ("4000-09-01T", 767255),
            ("4000-09-30T", 767284),
            ("4000-10-01T", 767285),
            ("4000-10-31T", 767315),
            ("4000-11-01T", 767316),
            ("4000-11-30T", 767345),
            ("4000-12-01T", 767346),
            ("4000-12-31T", 767376),
            ("4321-01-01T", 884254),
            ("4321-01-31T", 884284),
            ("4321-02-01T", 884285),
            ("4321-02-28T", 884312),
            ("4321-03-01T", 884313),
            ("4321-03-31T", 884343),
            ("4321-04-01T", 884344),
            ("4321-04-30T", 884373),
            ("4321-05-01T", 884374),
            ("4321-05-31T", 884404),
            ("4321-06-01T", 884405),
            ("4321-06-30T", 884434),
            ("4321-07-01T", 884435),
            ("4321-07-31T", 884465),
            ("4321-08-01T", 884466),
            ("4321-08-31T", 884496),
            ("4321-09-01T", 884497),
            ("4321-09-30T", 884526),
            ("4321-10-01T", 884527),
            ("4321-10-31T", 884557),
            ("4321-11-01T", 884558),
            ("4321-11-30T", 884587),
            ("4321-12-01T", 884588),
            ("4321-12-31T", 884618),
            ("9999-01-01T", 2958101),
            ("9999-01-31T", 2958131),
            ("9999-02-01T", 2958132),
            ("9999-02-28T", 2958159),
            ("9999-03-01T", 2958160),
            ("9999-03-31T", 2958190),
            ("9999-04-01T", 2958191),
            ("9999-04-30T", 2958220),
            ("9999-05-01T", 2958221),
            ("9999-05-31T", 2958251),
            ("9999-06-01T", 2958252),
            ("9999-06-30T", 2958281),
            ("9999-07-01T", 2958282),
            ("9999-07-31T", 2958312),
            ("9999-08-01T", 2958313),
            ("9999-08-31T", 2958343),
            ("9999-09-01T", 2958344),
            ("9999-09-30T", 2958373),
            ("9999-10-01T", 2958374),
            ("9999-10-31T", 2958404),
            ("9999-11-01T", 2958405),
            ("9999-11-30T", 2958434),
            ("9999-12-01T", 2958435),
            ("9999-12-31T", 2958465),
        ]

        # Dates and corresponding numbers from an Excel file.
        self.excel_1904_dates = [
            ("1904-01-01T", 0),
            ("1904-01-31T", 30),
            ("1904-02-01T", 31),
            ("1904-02-29T", 59),
            ("1904-03-01T", 60),
            ("1904-03-31T", 90),
            ("1904-04-01T", 91),
            ("1904-04-30T", 120),
            ("1904-05-01T", 121),
            ("1904-05-31T", 151),
            ("1904-06-01T", 152),
            ("1904-06-30T", 181),
            ("1904-07-01T", 182),
            ("1904-07-31T", 212),
            ("1904-08-01T", 213),
            ("1904-08-31T", 243),
            ("1904-09-01T", 244),
            ("1904-09-30T", 273),
            ("1904-10-01T", 274),
            ("1904-10-31T", 304),
            ("1904-11-01T", 305),
            ("1904-11-30T", 334),
            ("1904-12-01T", 335),
            ("1904-12-31T", 365),
            ("1907-02-27T", 1153),
            ("1907-02-28T", 1154),
            ("1907-03-01T", 1155),
            ("1907-03-02T", 1156),
            ("1907-03-03T", 1157),
            ("1907-03-04T", 1158),
            ("1907-03-05T", 1159),
            ("1907-03-06T", 1160),
            ("1999-01-01T", 34699),
            ("1999-01-31T", 34729),
            ("1999-02-01T", 34730),
            ("1999-02-28T", 34757),
            ("1999-03-01T", 34758),
            ("1999-03-31T", 34788),
            ("1999-04-01T", 34789),
            ("1999-04-30T", 34818),
            ("1999-05-01T", 34819),
            ("1999-05-31T", 34849),
            ("1999-06-01T", 34850),
            ("1999-06-30T", 34879),
            ("1999-07-01T", 34880),
            ("1999-07-31T", 34910),
            ("1999-08-01T", 34911),
            ("1999-08-31T", 34941),
            ("1999-09-01T", 34942),
            ("1999-09-30T", 34971),
            ("1999-10-01T", 34972),
            ("1999-10-31T", 35002),
            ("1999-11-01T", 35003),
            ("1999-11-30T", 35032),
            ("1999-12-01T", 35033),
            ("1999-12-31T", 35063),
            ("2000-01-01T", 35064),
            ("2000-01-31T", 35094),
            ("2000-02-01T", 35095),
            ("2000-02-29T", 35123),
            ("2000-03-01T", 35124),
            ("2000-03-31T", 35154),
            ("2000-04-01T", 35155),
            ("2000-04-30T", 35184),
            ("2000-05-01T", 35185),
            ("2000-05-31T", 35215),
            ("2000-06-01T", 35216),
            ("2000-06-30T", 35245),
            ("2000-07-01T", 35246),
            ("2000-07-31T", 35276),
            ("2000-08-01T", 35277),
            ("2000-08-31T", 35307),
            ("2000-09-01T", 35308),
            ("2000-09-30T", 35337),
            ("2000-10-01T", 35338),
            ("2000-10-31T", 35368),
            ("2000-11-01T", 35369),
            ("2000-11-30T", 35398),
            ("2000-12-01T", 35399),
            ("2000-12-31T", 35429),
            ("2001-01-01T", 35430),
            ("2001-01-31T", 35460),
            ("2001-02-01T", 35461),
            ("2001-02-28T", 35488),
            ("2001-03-01T", 35489),
            ("2001-03-31T", 35519),
            ("2001-04-01T", 35520),
            ("2001-04-30T", 35549),
            ("2001-05-01T", 35550),
            ("2001-05-31T", 35580),
            ("2001-06-01T", 35581),
            ("2001-06-30T", 35610),
            ("2001-07-01T", 35611),
            ("2001-07-31T", 35641),
            ("2001-08-01T", 35642),
            ("2001-08-31T", 35672),
            ("2001-09-01T", 35673),
            ("2001-09-30T", 35702),
            ("2001-10-01T", 35703),
            ("2001-10-31T", 35733),
            ("2001-11-01T", 35734),
            ("2001-11-30T", 35763),
            ("2001-12-01T", 35764),
            ("2001-12-31T", 35794),
            ("2400-01-01T", 181161),
            ("2400-01-31T", 181191),
            ("2400-02-01T", 181192),
            ("2400-02-29T", 181220),
            ("2400-03-01T", 181221),
            ("2400-03-31T", 181251),
            ("2400-04-01T", 181252),
            ("2400-04-30T", 181281),
            ("2400-05-01T", 181282),
            ("2400-05-31T", 181312),
            ("2400-06-01T", 181313),
            ("2400-06-30T", 181342),
            ("2400-07-01T", 181343),
            ("2400-07-31T", 181373),
            ("2400-08-01T", 181374),
            ("2400-08-31T", 181404),
            ("2400-09-01T", 181405),
            ("2400-09-30T", 181434),
            ("2400-10-01T", 181435),
            ("2400-10-31T", 181465),
            ("2400-11-01T", 181466),
            ("2400-11-30T", 181495),
            ("2400-12-01T", 181496),
            ("2400-12-31T", 181526),
            ("4000-01-01T", 765549),
            ("4000-01-31T", 765579),
            ("4000-02-01T", 765580),
            ("4000-02-29T", 765608),
            ("4000-03-01T", 765609),
            ("4000-03-31T", 765639),
            ("4000-04-01T", 765640),
            ("4000-04-30T", 765669),
            ("4000-05-01T", 765670),
            ("4000-05-31T", 765700),
            ("4000-06-01T", 765701),
            ("4000-06-30T", 765730),
            ("4000-07-01T", 765731),
            ("4000-07-31T", 765761),
            ("4000-08-01T", 765762),
            ("4000-08-31T", 765792),
            ("4000-09-01T", 765793),
            ("4000-09-30T", 765822),
            ("4000-10-01T", 765823),
            ("4000-10-31T", 765853),
            ("4000-11-01T", 765854),
            ("4000-11-30T", 765883),
            ("4000-12-01T", 765884),
            ("4000-12-31T", 765914),
            ("4321-01-01T", 882792),
            ("4321-01-31T", 882822),
            ("4321-02-01T", 882823),
            ("4321-02-28T", 882850),
            ("4321-03-01T", 882851),
            ("4321-03-31T", 882881),
            ("4321-04-01T", 882882),
            ("4321-04-30T", 882911),
            ("4321-05-01T", 882912),
            ("4321-05-31T", 882942),
            ("4321-06-01T", 882943),
            ("4321-06-30T", 882972),
            ("4321-07-01T", 882973),
            ("4321-07-31T", 883003),
            ("4321-08-01T", 883004),
            ("4321-08-31T", 883034),
            ("4321-09-01T", 883035),
            ("4321-09-30T", 883064),
            ("4321-10-01T", 883065),
            ("4321-10-31T", 883095),
            ("4321-11-01T", 883096),
            ("4321-11-30T", 883125),
            ("4321-12-01T", 883126),
            ("4321-12-31T", 883156),
            ("9999-01-01T", 2956639),
            ("9999-01-31T", 2956669),
            ("9999-02-01T", 2956670),
            ("9999-02-28T", 2956697),
            ("9999-03-01T", 2956698),
            ("9999-03-31T", 2956728),
            ("9999-04-01T", 2956729),
            ("9999-04-30T", 2956758),
            ("9999-05-01T", 2956759),
            ("9999-05-31T", 2956789),
            ("9999-06-01T", 2956790),
            ("9999-06-30T", 2956819),
            ("9999-07-01T", 2956820),
            ("9999-07-31T", 2956850),
            ("9999-08-01T", 2956851),
            ("9999-08-31T", 2956881),
            ("9999-09-01T", 2956882),
            ("9999-09-30T", 2956911),
            ("9999-10-01T", 2956912),
            ("9999-10-31T", 2956942),
            ("9999-11-01T", 2956943),
            ("9999-11-30T", 2956972),
            ("9999-12-01T", 2956973),
            ("9999-12-31T", 2957003),
        ]

    def test_convert_date_time_datetime(self):
        """Test the _convert_date_time() method with datetime objects."""

        for excel_date in self.excel_dates:
            test_date = datetime.datetime.strptime(excel_date[0], "%Y-%m-%dT")

            got = self.worksheet._convert_date_time(test_date)
            exp = excel_date[1]

            self.assertEqual(got, exp)

    def test_convert_date_time_date(self):
        """Test the _convert_date_time() method with date objects."""

        for excel_date in self.excel_dates:
            date_str = excel_date[0].rstrip("T")
            (year, month, day) = date_str.split("-")

            test_date = datetime.date(int(year), int(month), int(day))

            got = self.worksheet._convert_date_time(test_date)
            exp = excel_date[1]

            self.assertEqual(got, exp)

    def test_convert_date_time_1904(self):
        """Test the _convert_date_time() method with 1904 date system."""

        self.worksheet.date_1904 = True
        self.worksheet.epoch = datetime.datetime(1904, 1, 1)

        for excel_date in self.excel_1904_dates:
            date = datetime.datetime.strptime(excel_date[0], "%Y-%m-%dT")

            got = self.worksheet._convert_date_time(date)
            exp = excel_date[1]

            self.assertEqual(got, exp)
