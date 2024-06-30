###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from ...utility import quote_sheetname


class TestUtility(unittest.TestCase):
    """
    Test xl_cell_to_rowcol_abs() utility function.

    """

    def test_quote_sheetname(self):
        """Test xl_cell_to_rowcol_abs()"""

        # The following unquoted and quoted sheet names were extracted from
        # Excel files.
        tests = [
            # A sheetname that is already quoted.
            ("'Sheet 1'", "'Sheet 1'"),
            # ----------------------------------------------------------------
            # Rule 1.
            # ----------------------------------------------------------------
            # Some simple variants on standard sheet names.
            ("Sheet1", "Sheet1"),
            ("Sheet.1", "Sheet.1"),
            ("Sheet_1", "Sheet_1"),
            ("Sheet-1", "'Sheet-1'"),
            ("Sheet 1", "'Sheet 1'"),
            ("Sheet#1", "'Sheet#1'"),
            ("#Sheet1", "'#Sheet1'"),
            # Sheetnames with single quotes.
            ("Sheet'1", "'Sheet''1'"),
            ("Sheet''1", "'Sheet''''1'"),
            # Single special chars that are unquoted in sheetnames. These are
            # variants of the first char rule.
            ("_", "_"),
            (".", "'.'"),
            # White space only.
            (" ", "' '"),
            ("    ", "'    '"),
            # Sheetnames with unicode or emojis.
            ("été", "été"),
            ("mangé", "mangé"),
            ("Sheet😀", "Sheet😀"),
            ("Sheet🤌1", "Sheet🤌1"),
            ("Sheet⟦1", "'Sheet⟦1'"),  # Unicode punctuation.
            ("Sheet᠅1", "'Sheet᠅1'"),  # Unicode punctuation.
            # ----------------------------------------------------------------
            # Rule 2.
            # ----------------------------------------------------------------
            # Sheetnames starting with non-word characters.
            ("_Sheet1", "_Sheet1"),
            (".Sheet1", "'.Sheet1'"),
            ("1Sheet1", "'1Sheet1'"),
            ("-Sheet1", "'-Sheet1'"),
            ("😀Sheet", "'😀Sheet'"),
            # Sheetnames that are digits only also start with a non word char.
            ("1", "'1'"),
            ("2", "'2'"),
            ("1234", "'1234'"),
            ("12345678", "'12345678'"),
            # ----------------------------------------------------------------
            # Rule 3.
            # ----------------------------------------------------------------
            # Worksheet names that look like A1 style references (with the
            # row/column number in the Excel allowable range). These are case
            # insensitive.
            ("A0", "A0"),
            ("A1", "'A1'"),
            ("a1", "'a1'"),
            ("XFD", "XFD"),
            ("xfd", "xfd"),
            ("XFE1", "XFE1"),
            ("ZZZ1", "ZZZ1"),
            ("XFD1", "'XFD1'"),
            ("xfd1", "'xfd1'"),
            ("B1048577", "B1048577"),
            ("A1048577", "A1048577"),
            ("A1048576", "'A1048576'"),
            ("B1048576", "'B1048576'"),
            ("B1048576a", "B1048576a"),
            ("XFD048576", "'XFD048576'"),
            ("XFD1048576", "'XFD1048576'"),
            ("XFD01048577", "XFD01048577"),
            ("XFD01048576", "'XFD01048576'"),
            ("A123456789012345678901", "A123456789012345678901"),  # Exceeds u64.
            # ----------------------------------------------------------------
            # Rule 4.
            # ----------------------------------------------------------------
            # Sheet names that *start* with RC style references (with the
            # row/column number in the Excel allowable range). These are case
            # insensitive.
            ("A", "A"),
            ("B", "B"),
            ("D", "D"),
            ("Q", "Q"),
            ("S", "S"),
            ("c", "'c'"),
            ("C", "'C'"),
            ("CR", "CR"),
            ("CZ", "CZ"),
            ("r", "'r'"),
            ("R", "'R'"),
            ("C8", "'C8'"),
            ("rc", "'rc'"),
            ("RC", "'RC'"),
            ("RCZ", "RCZ"),
            ("RRC", "RRC"),
            ("R0C0", "R0C0"),
            ("R4C", "'R4C'"),
            ("R5C", "'R5C'"),
            ("rc2", "'rc2'"),
            ("RC2", "'RC2'"),
            ("RC8", "'RC8'"),
            ("bR1C1", "bR1C1"),
            ("R1C1", "'R1C1'"),
            ("r1c2", "'r1c2'"),
            ("rc2z", "'rc2z'"),
            ("bR1C1b", "bR1C1b"),
            ("R1C1b", "'R1C1b'"),
            ("R1C1R", "'R1C1R'"),
            ("C16384", "'C16384'"),
            ("C16385", "'C16385'"),
            ("C16385Z", "C16385Z"),
            ("C16386", "'C16386'"),
            ("C16384Z", "'C16384Z'"),
            ("PC16384Z", "PC16384Z"),
            ("RC16383", "'RC16383'"),
            ("RC16385Z", "RC16385Z"),
            ("R1048576", "'R1048576'"),
            ("R1048577C", "R1048577C"),
            ("R1C16384", "'R1C16384'"),
            ("R1C16385", "'R1C16385'"),
            ("RC16384Z", "'RC16384Z'"),
            ("R1048576C", "'R1048576C'"),
            ("R1048577C1", "R1048577C1"),
            ("R1C16384Z", "'R1C16384Z'"),
            ("R1048575C1", "'R1048575C1'"),
            ("R1048576C1", "'R1048576C1'"),
            ("R1048577C16384", "R1048577C16384"),
            ("R1048576C16384", "'R1048576C16384'"),
            ("R1048576C16385", "'R1048576C16385'"),
            ("ZR1048576C16384", "ZR1048576C16384"),
            ("C123456789012345678901Z", "C123456789012345678901Z"),  # Exceeds u64.
            ("R123456789012345678901Z", "R123456789012345678901Z"),  # Exceeds u64.
        ]

        for sheetname, exp in tests:
            got = quote_sheetname(sheetname)
            self.assertEqual(got, exp)
