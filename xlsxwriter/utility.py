###############################################################################
#
# Worksheet - A class for writing Excel Worksheets.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2025, John McNamara, jmcnamara@cpan.org
#
import datetime
import hashlib
import os
import re
from io import BytesIO
from struct import unpack
from typing import Dict, Optional, Tuple, Union
from warnings import warn

from .exceptions import UndefinedImageSize, UnsupportedImageFormat

COL_NAMES: Dict[int, str] = {}

CHAR_WIDTHS = {
    " ": 3,
    "!": 5,
    '"': 6,
    "#": 7,
    "$": 7,
    "%": 11,
    "&": 10,
    "'": 3,
    "(": 5,
    ")": 5,
    "*": 7,
    "+": 7,
    ",": 4,
    "-": 5,
    ".": 4,
    "/": 6,
    "0": 7,
    "1": 7,
    "2": 7,
    "3": 7,
    "4": 7,
    "5": 7,
    "6": 7,
    "7": 7,
    "8": 7,
    "9": 7,
    ":": 4,
    ";": 4,
    "<": 7,
    "=": 7,
    ">": 7,
    "?": 7,
    "@": 13,
    "A": 9,
    "B": 8,
    "C": 8,
    "D": 9,
    "E": 7,
    "F": 7,
    "G": 9,
    "H": 9,
    "I": 4,
    "J": 5,
    "K": 8,
    "L": 6,
    "M": 12,
    "N": 10,
    "O": 10,
    "P": 8,
    "Q": 10,
    "R": 8,
    "S": 7,
    "T": 7,
    "U": 9,
    "V": 9,
    "W": 13,
    "X": 8,
    "Y": 7,
    "Z": 7,
    "[": 5,
    "\\": 6,
    "]": 5,
    "^": 7,
    "_": 7,
    "`": 4,
    "a": 7,
    "b": 8,
    "c": 6,
    "d": 8,
    "e": 8,
    "f": 5,
    "g": 7,
    "h": 8,
    "i": 4,
    "j": 4,
    "k": 7,
    "l": 4,
    "m": 12,
    "n": 8,
    "o": 8,
    "p": 8,
    "q": 8,
    "r": 5,
    "s": 6,
    "t": 5,
    "u": 8,
    "v": 7,
    "w": 11,
    "x": 7,
    "y": 7,
    "z": 6,
    "{": 5,
    "|": 7,
    "}": 5,
    "~": 7,
}

# The following is a list of Emojis used to decide if worksheet names require
# quoting since there is (currently) no native support for matching them in
# Python regular expressions. It is probably unnecessary to exclude them since
# the default quoting is safe in Excel even when unnecessary (the reverse isn't
# true). The Emoji list was generated from:
#
# https://util.unicode.org/UnicodeJsps/list-unicodeset.jsp?a=%5B%3AEmoji%3DYes%3A%5D&abb=on&esc=on&g=&i=
#
# pylint: disable-next=line-too-long
EMOJIS = "\u00a9\u00ae\u203c\u2049\u2122\u2139\u2194-\u2199\u21a9\u21aa\u231a\u231b\u2328\u23cf\u23e9-\u23f3\u23f8-\u23fa\u24c2\u25aa\u25ab\u25b6\u25c0\u25fb-\u25fe\u2600-\u2604\u260e\u2611\u2614\u2615\u2618\u261d\u2620\u2622\u2623\u2626\u262a\u262e\u262f\u2638-\u263a\u2640\u2642\u2648-\u2653\u265f\u2660\u2663\u2665\u2666\u2668\u267b\u267e\u267f\u2692-\u2697\u2699\u269b\u269c\u26a0\u26a1\u26a7\u26aa\u26ab\u26b0\u26b1\u26bd\u26be\u26c4\u26c5\u26c8\u26ce\u26cf\u26d1\u26d3\u26d4\u26e9\u26ea\u26f0-\u26f5\u26f7-\u26fa\u26fd\u2702\u2705\u2708-\u270d\u270f\u2712\u2714\u2716\u271d\u2721\u2728\u2733\u2734\u2744\u2747\u274c\u274e\u2753-\u2755\u2757\u2763\u2764\u2795-\u2797\u27a1\u27b0\u27bf\u2934\u2935\u2b05-\u2b07\u2b1b\u2b1c\u2b50\u2b55\u3030\u303d\u3297\u3299\U0001f004\U0001f0cf\U0001f170\U0001f171\U0001f17e\U0001f17f\U0001f18e\U0001f191-\U0001f19a\U0001f1e6-\U0001f1ff\U0001f201\U0001f202\U0001f21a\U0001f22f\U0001f232-\U0001f23a\U0001f250\U0001f251\U0001f300-\U0001f321\U0001f324-\U0001f393\U0001f396\U0001f397\U0001f399-\U0001f39b\U0001f39e-\U0001f3f0\U0001f3f3-\U0001f3f5\U0001f3f7-\U0001f4fd\U0001f4ff-\U0001f53d\U0001f549-\U0001f54e\U0001f550-\U0001f567\U0001f56f\U0001f570\U0001f573-\U0001f57a\U0001f587\U0001f58a-\U0001f58d\U0001f590\U0001f595\U0001f596\U0001f5a4\U0001f5a5\U0001f5a8\U0001f5b1\U0001f5b2\U0001f5bc\U0001f5c2-\U0001f5c4\U0001f5d1-\U0001f5d3\U0001f5dc-\U0001f5de\U0001f5e1\U0001f5e3\U0001f5e8\U0001f5ef\U0001f5f3\U0001f5fa-\U0001f64f\U0001f680-\U0001f6c5\U0001f6cb-\U0001f6d2\U0001f6d5-\U0001f6d7\U0001f6dc-\U0001f6e5\U0001f6e9\U0001f6eb\U0001f6ec\U0001f6f0\U0001f6f3-\U0001f6fc\U0001f7e0-\U0001f7eb\U0001f7f0\U0001f90c-\U0001f93a\U0001f93c-\U0001f945\U0001f947-\U0001f9ff\U0001fa70-\U0001fa7c\U0001fa80-\U0001fa88\U0001fa90-\U0001fabd\U0001fabf-\U0001fac5\U0001face-\U0001fadb\U0001fae0-\U0001fae8\U0001faf0-\U0001faf8"  # noqa

# Compile performance critical regular expressions.
RE_LEADING_WHITESPACE = re.compile(r"^\s")
RE_TRAILING_WHITESPACE = re.compile(r"\s$")
RE_RANGE_PARTS = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")
RE_QUOTE_RULE1 = re.compile(rf"[^\w\.{EMOJIS}]")
RE_QUOTE_RULE2 = re.compile(rf"^[\d\.{EMOJIS}]")
RE_QUOTE_RULE3 = re.compile(r"^([A-Z]{1,3}\d+)$")
RE_QUOTE_RULE4_ROW = re.compile(r"^R(\d+)")
RE_QUOTE_RULE4_COLUMN = re.compile(r"^R?C(\d+)")


def xl_rowcol_to_cell(
    row: int,
    col: int,
    row_abs: bool = False,
    col_abs: bool = False,
) -> str:
    """
    Convert a zero indexed row and column cell reference to a A1 style string.

    Args:
       row:     The cell row.    Int.
       col:     The cell column. Int.
       row_abs: Optional flag to make the row absolute.    Bool.
       col_abs: Optional flag to make the column absolute. Bool.

    Returns:
        A1 style string.

    """
    if row < 0:
        warn(f"Row number '{row}' must be >= 0")
        return ""

    if col < 0:
        warn(f"Col number '{col}' must be >= 0")
        return ""

    row += 1  # Change to 1-index.
    row_abs_str = "$" if row_abs else ""

    col_str = xl_col_to_name(col, col_abs)

    return col_str + row_abs_str + str(row)


def xl_rowcol_to_cell_fast(row: int, col: int) -> str:
    """
    Optimized version of the xl_rowcol_to_cell function. Only used internally.

    Args:
       row: The cell row.    Int.
       col: The cell column. Int.

    Returns:
        A1 style string.

    """
    if col in COL_NAMES:
        col_str = COL_NAMES[col]
    else:
        col_str = xl_col_to_name(col)
        COL_NAMES[col] = col_str

    return col_str + str(row + 1)


def xl_col_to_name(col: int, col_abs: bool = False) -> str:
    """
    Convert a zero indexed column cell reference to a string.

    Args:
       col:     The cell column. Int.
       col_abs: Optional flag to make the column absolute. Bool.

    Returns:
        Column style string.

    """
    col_num = col
    if col_num < 0:
        warn(f"Col number '{col_num}' must be >= 0")
        return ""

    col_num += 1  # Change to 1-index.
    col_str = ""
    col_abs_str = "$" if col_abs else ""

    while col_num:
        # Set remainder from 1 .. 26
        remainder = col_num % 26

        if remainder == 0:
            remainder = 26

        # Convert the remainder to a character.
        col_letter = chr(ord("A") + remainder - 1)

        # Accumulate the column letters, right to left.
        col_str = col_letter + col_str

        # Get the next order of magnitude.
        col_num = int((col_num - 1) / 26)

    return col_abs_str + col_str


def xl_cell_to_rowcol(cell_str: str) -> Tuple[int, int]:
    """
    Convert a cell reference in A1 notation to a zero indexed row and column.

    Args:
       cell_str:  A1 style string.

    Returns:
        row, col: Zero indexed cell row and column indices.

    """
    if not cell_str:
        return 0, 0

    match = RE_RANGE_PARTS.match(cell_str)
    if match is None:
        warn(f"Invalid cell reference '{cell_str}'")
        return 0, 0

    col_str = match.group(2)
    row_str = match.group(4)

    # Convert base26 column string to number.
    expn = 0
    col = 0
    for char in reversed(col_str):
        col += (ord(char) - ord("A") + 1) * (26**expn)
        expn += 1

    # Convert 1-index to zero-index
    row = int(row_str) - 1
    col -= 1

    return row, col


def xl_cell_to_rowcol_abs(cell_str: str) -> Tuple[int, int, bool, bool]:
    """
    Convert an absolute cell reference in A1 notation to a zero indexed
    row and column, with True/False values for absolute rows or columns.

    Args:
       cell_str: A1 style string.

    Returns:
        row, col, row_abs, col_abs:  Zero indexed cell row and column indices.

    """
    if not cell_str:
        return 0, 0, False, False

    match = RE_RANGE_PARTS.match(cell_str)
    if match is None:
        warn(f"Invalid cell reference '{cell_str}'")
        return 0, 0, False, False

    col_abs = bool(match.group(1))
    col_str = match.group(2)
    row_abs = bool(match.group(3))
    row_str = match.group(4)

    # Convert base26 column string to number.
    expn = 0
    col = 0
    for char in reversed(col_str):
        col += (ord(char) - ord("A") + 1) * (26**expn)
        expn += 1

    # Convert 1-index to zero-index
    row = int(row_str) - 1
    col -= 1

    return row, col, row_abs, col_abs


def xl_range(first_row: int, first_col: int, last_row: int, last_col: int) -> str:
    """
    Convert zero indexed row and col cell references to a A1:B1 range string.

    Args:
       first_row: The first cell row.    Int.
       first_col: The first cell column. Int.
       last_row:  The last cell row.     Int.
       last_col:  The last cell column.  Int.

    Returns:
        A1:B1 style range string.

    """
    range1 = xl_rowcol_to_cell(first_row, first_col)
    range2 = xl_rowcol_to_cell(last_row, last_col)

    if range1 == "" or range2 == "":
        warn("Row and column numbers must be >= 0")
        return ""

    if range1 == range2:
        return range1

    return range1 + ":" + range2


def xl_range_abs(first_row: int, first_col: int, last_row: int, last_col: int) -> str:
    """
    Convert zero indexed row and col cell references to a $A$1:$B$1 absolute
    range string.

    Args:
       first_row: The first cell row.    Int.
       first_col: The first cell column. Int.
       last_row:  The last cell row.     Int.
       last_col:  The last cell column.  Int.

    Returns:
        $A$1:$B$1 style range string.

    """
    range1 = xl_rowcol_to_cell(first_row, first_col, True, True)
    range2 = xl_rowcol_to_cell(last_row, last_col, True, True)

    if range1 == "" or range2 == "":
        warn("Row and column numbers must be >= 0")
        return ""

    if range1 == range2:
        return range1

    return range1 + ":" + range2


def xl_range_formula(
    sheetname: str, first_row: int, first_col: int, last_row: int, last_col: int
) -> str:
    """
    Convert worksheet name and zero indexed row and col cell references to
    a Sheet1!A1:B1 range formula string.

    Args:
       sheetname: The worksheet name.    String.
       first_row: The first cell row.    Int.
       first_col: The first cell column. Int.
       last_row:  The last cell row.     Int.
       last_col:  The last cell column.  Int.

    Returns:
        A1:B1 style range string.

    """
    cell_range = xl_range_abs(first_row, first_col, last_row, last_col)
    sheetname = quote_sheetname(sheetname)

    return sheetname + "!" + cell_range


def quote_sheetname(sheetname: str) -> str:
    """
    Sheetnames used in references should be quoted if they contain any spaces,
    special characters or if they look like a A1 or RC cell reference. The rules
    are shown inline below.

    Args:
       sheetname: The worksheet name. String.

    Returns:
        A quoted worksheet string.

    """
    uppercase_sheetname = sheetname.upper()
    requires_quoting = False
    col_max = 163_84
    row_max = 1048576

    # Don't quote sheetname if it is already quoted by the user.
    if not sheetname.startswith("'"):

        match_rule3 = RE_QUOTE_RULE3.match(uppercase_sheetname)
        match_rule4_row = RE_QUOTE_RULE4_ROW.match(uppercase_sheetname)
        match_rule4_column = RE_QUOTE_RULE4_COLUMN.match(uppercase_sheetname)

        # --------------------------------------------------------------------
        # Rule 1. Sheet names that contain anything other than \w and "."
        # characters must be quoted.
        # --------------------------------------------------------------------
        if RE_QUOTE_RULE1.search(sheetname):
            requires_quoting = True

        # --------------------------------------------------------------------
        # Rule 2. Sheet names that start with a digit or "." must be quoted.
        # --------------------------------------------------------------------
        elif RE_QUOTE_RULE2.search(sheetname):
            requires_quoting = True

        # --------------------------------------------------------------------
        # Rule 3. Sheet names must not be a valid A1 style cell reference.
        # Valid means that the row and column range values must also be within
        # Excel row and column limits.
        # --------------------------------------------------------------------
        elif match_rule3:
            cell = match_rule3.group(1)
            (row, col) = xl_cell_to_rowcol(cell)

            if 0 <= row < row_max and 0 <= col < col_max:
                requires_quoting = True

        # --------------------------------------------------------------------
        # Rule 4. Sheet names must not *start* with a valid RC style cell
        # reference. Other characters after the valid RC reference are ignored
        # by Excel. Valid means that the row and column range values must also
        # be within Excel row and column limits.
        #
        # Note: references without trailing characters like R12345 or C12345
        # are caught by Rule 3. Negative references like R-12345 are caught by
        # Rule 1 due to the dash.
        # --------------------------------------------------------------------

        # Rule 4a. Check for sheet names that start with R1 style references.
        elif match_rule4_row:
            row = int(match_rule4_row.group(1))

            if 0 < row <= row_max:
                requires_quoting = True

        # Rule 4b. Check for sheet names that start with C1 or RC1 style
        elif match_rule4_column:
            col = int(match_rule4_column.group(1))

            if 0 < col <= col_max:
                requires_quoting = True

        # Rule 4c. Check for some single R/C references.
        elif uppercase_sheetname in ("R", "C", "RC"):
            requires_quoting = True

    if requires_quoting:
        # Double quote any single quotes.
        sheetname = sheetname.replace("'", "''")

        # Single quote the sheet name.
        sheetname = f"'{sheetname}'"

    return sheetname


def cell_autofit_width(string: str) -> int:
    """
    Calculate the width required to auto-fit a string in a cell.

    Args:
       string: The string to calculate the cell width for. String.

    Returns:
        The string autofit width in pixels. Returns 0 if the string is empty.

    """
    if not string or len(string) == 0:
        return 0

    # Excel adds an additional 7 pixels of padding to the cell boundary.
    return xl_pixel_width(string) + 7


def xl_pixel_width(string: str) -> int:
    """
    Get the pixel width of a string based on individual character widths taken
    from Excel. UTF8 characters, and other unhandled characters, are given a
    default width of 8.

    Args:
       string: The string to calculate the width for. String.

    Returns:
        The string width in pixels. Note, Excel adds an additional 7 pixels of
        padding in the cell.

    """
    length = 0
    for char in string:
        length += CHAR_WIDTHS.get(char, 8)

    return length


def _xl_color(color: str) -> str:
    # Used in conjunction with the XlsxWriter *color() methods to convert
    # a color name into an RGB formatted string. These colors are for
    # backward compatibility with older versions of Excel.
    named_colors = {
        "black": "#000000",
        "blue": "#0000FF",
        "brown": "#800000",
        "cyan": "#00FFFF",
        "gray": "#808080",
        "green": "#008000",
        "lime": "#00FF00",
        "magenta": "#FF00FF",
        "navy": "#000080",
        "orange": "#FF6600",
        "pink": "#FF00FF",
        "purple": "#800080",
        "red": "#FF0000",
        "silver": "#C0C0C0",
        "white": "#FFFFFF",
        "yellow": "#FFFF00",
    }

    color = named_colors.get(color, color)

    if not re.match("#[0-9a-fA-F]{6}", color):
        warn(f"Color '{color}' isn't a valid Excel color")

    # Convert the RGB color to the Excel ARGB format.
    return "FF" + color.lstrip("#").upper()


def _get_rgb_color(color: str) -> str:
    # Convert the user specified color to an RGB color.
    rgb_color = _xl_color(color)

    # Remove leading FF from RGB color for charts.
    rgb_color = re.sub(r"^FF", "", rgb_color)

    return rgb_color


def _get_sparkline_style(style_id: int) -> Dict[str, Dict[str, str]]:
    styles = [
        {
            "series": {"theme": "4", "tint": "-0.499984740745262"},
            "negative": {"theme": "5"},
            "markers": {"theme": "4", "tint": "-0.499984740745262"},
            "first": {"theme": "4", "tint": "0.39997558519241921"},
            "last": {"theme": "4", "tint": "0.39997558519241921"},
            "high": {"theme": "4"},
            "low": {"theme": "4"},
        },  # 0
        {
            "series": {"theme": "4", "tint": "-0.499984740745262"},
            "negative": {"theme": "5"},
            "markers": {"theme": "4", "tint": "-0.499984740745262"},
            "first": {"theme": "4", "tint": "0.39997558519241921"},
            "last": {"theme": "4", "tint": "0.39997558519241921"},
            "high": {"theme": "4"},
            "low": {"theme": "4"},
        },  # 1
        {
            "series": {"theme": "5", "tint": "-0.499984740745262"},
            "negative": {"theme": "6"},
            "markers": {"theme": "5", "tint": "-0.499984740745262"},
            "first": {"theme": "5", "tint": "0.39997558519241921"},
            "last": {"theme": "5", "tint": "0.39997558519241921"},
            "high": {"theme": "5"},
            "low": {"theme": "5"},
        },  # 2
        {
            "series": {"theme": "6", "tint": "-0.499984740745262"},
            "negative": {"theme": "7"},
            "markers": {"theme": "6", "tint": "-0.499984740745262"},
            "first": {"theme": "6", "tint": "0.39997558519241921"},
            "last": {"theme": "6", "tint": "0.39997558519241921"},
            "high": {"theme": "6"},
            "low": {"theme": "6"},
        },  # 3
        {
            "series": {"theme": "7", "tint": "-0.499984740745262"},
            "negative": {"theme": "8"},
            "markers": {"theme": "7", "tint": "-0.499984740745262"},
            "first": {"theme": "7", "tint": "0.39997558519241921"},
            "last": {"theme": "7", "tint": "0.39997558519241921"},
            "high": {"theme": "7"},
            "low": {"theme": "7"},
        },  # 4
        {
            "series": {"theme": "8", "tint": "-0.499984740745262"},
            "negative": {"theme": "9"},
            "markers": {"theme": "8", "tint": "-0.499984740745262"},
            "first": {"theme": "8", "tint": "0.39997558519241921"},
            "last": {"theme": "8", "tint": "0.39997558519241921"},
            "high": {"theme": "8"},
            "low": {"theme": "8"},
        },  # 5
        {
            "series": {"theme": "9", "tint": "-0.499984740745262"},
            "negative": {"theme": "4"},
            "markers": {"theme": "9", "tint": "-0.499984740745262"},
            "first": {"theme": "9", "tint": "0.39997558519241921"},
            "last": {"theme": "9", "tint": "0.39997558519241921"},
            "high": {"theme": "9"},
            "low": {"theme": "9"},
        },  # 6
        {
            "series": {"theme": "4", "tint": "-0.249977111117893"},
            "negative": {"theme": "5"},
            "markers": {"theme": "5", "tint": "-0.249977111117893"},
            "first": {"theme": "5", "tint": "-0.249977111117893"},
            "last": {"theme": "5", "tint": "-0.249977111117893"},
            "high": {"theme": "5", "tint": "-0.249977111117893"},
            "low": {"theme": "5", "tint": "-0.249977111117893"},
        },  # 7
        {
            "series": {"theme": "5", "tint": "-0.249977111117893"},
            "negative": {"theme": "6"},
            "markers": {"theme": "6", "tint": "-0.249977111117893"},
            "first": {"theme": "6", "tint": "-0.249977111117893"},
            "last": {"theme": "6", "tint": "-0.249977111117893"},
            "high": {"theme": "6", "tint": "-0.249977111117893"},
            "low": {"theme": "6", "tint": "-0.249977111117893"},
        },  # 8
        {
            "series": {"theme": "6", "tint": "-0.249977111117893"},
            "negative": {"theme": "7"},
            "markers": {"theme": "7", "tint": "-0.249977111117893"},
            "first": {"theme": "7", "tint": "-0.249977111117893"},
            "last": {"theme": "7", "tint": "-0.249977111117893"},
            "high": {"theme": "7", "tint": "-0.249977111117893"},
            "low": {"theme": "7", "tint": "-0.249977111117893"},
        },  # 9
        {
            "series": {"theme": "7", "tint": "-0.249977111117893"},
            "negative": {"theme": "8"},
            "markers": {"theme": "8", "tint": "-0.249977111117893"},
            "first": {"theme": "8", "tint": "-0.249977111117893"},
            "last": {"theme": "8", "tint": "-0.249977111117893"},
            "high": {"theme": "8", "tint": "-0.249977111117893"},
            "low": {"theme": "8", "tint": "-0.249977111117893"},
        },  # 10
        {
            "series": {"theme": "8", "tint": "-0.249977111117893"},
            "negative": {"theme": "9"},
            "markers": {"theme": "9", "tint": "-0.249977111117893"},
            "first": {"theme": "9", "tint": "-0.249977111117893"},
            "last": {"theme": "9", "tint": "-0.249977111117893"},
            "high": {"theme": "9", "tint": "-0.249977111117893"},
            "low": {"theme": "9", "tint": "-0.249977111117893"},
        },  # 11
        {
            "series": {"theme": "9", "tint": "-0.249977111117893"},
            "negative": {"theme": "4"},
            "markers": {"theme": "4", "tint": "-0.249977111117893"},
            "first": {"theme": "4", "tint": "-0.249977111117893"},
            "last": {"theme": "4", "tint": "-0.249977111117893"},
            "high": {"theme": "4", "tint": "-0.249977111117893"},
            "low": {"theme": "4", "tint": "-0.249977111117893"},
        },  # 12
        {
            "series": {"theme": "4"},
            "negative": {"theme": "5"},
            "markers": {"theme": "4", "tint": "-0.249977111117893"},
            "first": {"theme": "4", "tint": "-0.249977111117893"},
            "last": {"theme": "4", "tint": "-0.249977111117893"},
            "high": {"theme": "4", "tint": "-0.249977111117893"},
            "low": {"theme": "4", "tint": "-0.249977111117893"},
        },  # 13
        {
            "series": {"theme": "5"},
            "negative": {"theme": "6"},
            "markers": {"theme": "5", "tint": "-0.249977111117893"},
            "first": {"theme": "5", "tint": "-0.249977111117893"},
            "last": {"theme": "5", "tint": "-0.249977111117893"},
            "high": {"theme": "5", "tint": "-0.249977111117893"},
            "low": {"theme": "5", "tint": "-0.249977111117893"},
        },  # 14
        {
            "series": {"theme": "6"},
            "negative": {"theme": "7"},
            "markers": {"theme": "6", "tint": "-0.249977111117893"},
            "first": {"theme": "6", "tint": "-0.249977111117893"},
            "last": {"theme": "6", "tint": "-0.249977111117893"},
            "high": {"theme": "6", "tint": "-0.249977111117893"},
            "low": {"theme": "6", "tint": "-0.249977111117893"},
        },  # 15
        {
            "series": {"theme": "7"},
            "negative": {"theme": "8"},
            "markers": {"theme": "7", "tint": "-0.249977111117893"},
            "first": {"theme": "7", "tint": "-0.249977111117893"},
            "last": {"theme": "7", "tint": "-0.249977111117893"},
            "high": {"theme": "7", "tint": "-0.249977111117893"},
            "low": {"theme": "7", "tint": "-0.249977111117893"},
        },  # 16
        {
            "series": {"theme": "8"},
            "negative": {"theme": "9"},
            "markers": {"theme": "8", "tint": "-0.249977111117893"},
            "first": {"theme": "8", "tint": "-0.249977111117893"},
            "last": {"theme": "8", "tint": "-0.249977111117893"},
            "high": {"theme": "8", "tint": "-0.249977111117893"},
            "low": {"theme": "8", "tint": "-0.249977111117893"},
        },  # 17
        {
            "series": {"theme": "9"},
            "negative": {"theme": "4"},
            "markers": {"theme": "9", "tint": "-0.249977111117893"},
            "first": {"theme": "9", "tint": "-0.249977111117893"},
            "last": {"theme": "9", "tint": "-0.249977111117893"},
            "high": {"theme": "9", "tint": "-0.249977111117893"},
            "low": {"theme": "9", "tint": "-0.249977111117893"},
        },  # 18
        {
            "series": {"theme": "4", "tint": "0.39997558519241921"},
            "negative": {"theme": "0", "tint": "-0.499984740745262"},
            "markers": {"theme": "4", "tint": "0.79998168889431442"},
            "first": {"theme": "4", "tint": "-0.249977111117893"},
            "last": {"theme": "4", "tint": "-0.249977111117893"},
            "high": {"theme": "4", "tint": "-0.499984740745262"},
            "low": {"theme": "4", "tint": "-0.499984740745262"},
        },  # 19
        {
            "series": {"theme": "5", "tint": "0.39997558519241921"},
            "negative": {"theme": "0", "tint": "-0.499984740745262"},
            "markers": {"theme": "5", "tint": "0.79998168889431442"},
            "first": {"theme": "5", "tint": "-0.249977111117893"},
            "last": {"theme": "5", "tint": "-0.249977111117893"},
            "high": {"theme": "5", "tint": "-0.499984740745262"},
            "low": {"theme": "5", "tint": "-0.499984740745262"},
        },  # 20
        {
            "series": {"theme": "6", "tint": "0.39997558519241921"},
            "negative": {"theme": "0", "tint": "-0.499984740745262"},
            "markers": {"theme": "6", "tint": "0.79998168889431442"},
            "first": {"theme": "6", "tint": "-0.249977111117893"},
            "last": {"theme": "6", "tint": "-0.249977111117893"},
            "high": {"theme": "6", "tint": "-0.499984740745262"},
            "low": {"theme": "6", "tint": "-0.499984740745262"},
        },  # 21
        {
            "series": {"theme": "7", "tint": "0.39997558519241921"},
            "negative": {"theme": "0", "tint": "-0.499984740745262"},
            "markers": {"theme": "7", "tint": "0.79998168889431442"},
            "first": {"theme": "7", "tint": "-0.249977111117893"},
            "last": {"theme": "7", "tint": "-0.249977111117893"},
            "high": {"theme": "7", "tint": "-0.499984740745262"},
            "low": {"theme": "7", "tint": "-0.499984740745262"},
        },  # 22
        {
            "series": {"theme": "8", "tint": "0.39997558519241921"},
            "negative": {"theme": "0", "tint": "-0.499984740745262"},
            "markers": {"theme": "8", "tint": "0.79998168889431442"},
            "first": {"theme": "8", "tint": "-0.249977111117893"},
            "last": {"theme": "8", "tint": "-0.249977111117893"},
            "high": {"theme": "8", "tint": "-0.499984740745262"},
            "low": {"theme": "8", "tint": "-0.499984740745262"},
        },  # 23
        {
            "series": {"theme": "9", "tint": "0.39997558519241921"},
            "negative": {"theme": "0", "tint": "-0.499984740745262"},
            "markers": {"theme": "9", "tint": "0.79998168889431442"},
            "first": {"theme": "9", "tint": "-0.249977111117893"},
            "last": {"theme": "9", "tint": "-0.249977111117893"},
            "high": {"theme": "9", "tint": "-0.499984740745262"},
            "low": {"theme": "9", "tint": "-0.499984740745262"},
        },  # 24
        {
            "series": {"theme": "1", "tint": "0.499984740745262"},
            "negative": {"theme": "1", "tint": "0.249977111117893"},
            "markers": {"theme": "1", "tint": "0.249977111117893"},
            "first": {"theme": "1", "tint": "0.249977111117893"},
            "last": {"theme": "1", "tint": "0.249977111117893"},
            "high": {"theme": "1", "tint": "0.249977111117893"},
            "low": {"theme": "1", "tint": "0.249977111117893"},
        },  # 25
        {
            "series": {"theme": "1", "tint": "0.34998626667073579"},
            "negative": {"theme": "0", "tint": "-0.249977111117893"},
            "markers": {"theme": "0", "tint": "-0.249977111117893"},
            "first": {"theme": "0", "tint": "-0.249977111117893"},
            "last": {"theme": "0", "tint": "-0.249977111117893"},
            "high": {"theme": "0", "tint": "-0.249977111117893"},
            "low": {"theme": "0", "tint": "-0.249977111117893"},
        },  # 26
        {
            "series": {"rgb": "FF323232"},
            "negative": {"rgb": "FFD00000"},
            "markers": {"rgb": "FFD00000"},
            "first": {"rgb": "FFD00000"},
            "last": {"rgb": "FFD00000"},
            "high": {"rgb": "FFD00000"},
            "low": {"rgb": "FFD00000"},
        },  # 27
        {
            "series": {"rgb": "FF000000"},
            "negative": {"rgb": "FF0070C0"},
            "markers": {"rgb": "FF0070C0"},
            "first": {"rgb": "FF0070C0"},
            "last": {"rgb": "FF0070C0"},
            "high": {"rgb": "FF0070C0"},
            "low": {"rgb": "FF0070C0"},
        },  # 28
        {
            "series": {"rgb": "FF376092"},
            "negative": {"rgb": "FFD00000"},
            "markers": {"rgb": "FFD00000"},
            "first": {"rgb": "FFD00000"},
            "last": {"rgb": "FFD00000"},
            "high": {"rgb": "FFD00000"},
            "low": {"rgb": "FFD00000"},
        },  # 29
        {
            "series": {"rgb": "FF0070C0"},
            "negative": {"rgb": "FF000000"},
            "markers": {"rgb": "FF000000"},
            "first": {"rgb": "FF000000"},
            "last": {"rgb": "FF000000"},
            "high": {"rgb": "FF000000"},
            "low": {"rgb": "FF000000"},
        },  # 30
        {
            "series": {"rgb": "FF5F5F5F"},
            "negative": {"rgb": "FFFFB620"},
            "markers": {"rgb": "FFD70077"},
            "first": {"rgb": "FF5687C2"},
            "last": {"rgb": "FF359CEB"},
            "high": {"rgb": "FF56BE79"},
            "low": {"rgb": "FFFF5055"},
        },  # 31
        {
            "series": {"rgb": "FF5687C2"},
            "negative": {"rgb": "FFFFB620"},
            "markers": {"rgb": "FFD70077"},
            "first": {"rgb": "FF777777"},
            "last": {"rgb": "FF359CEB"},
            "high": {"rgb": "FF56BE79"},
            "low": {"rgb": "FFFF5055"},
        },  # 32
        {
            "series": {"rgb": "FFC6EFCE"},
            "negative": {"rgb": "FFFFC7CE"},
            "markers": {"rgb": "FF8CADD6"},
            "first": {"rgb": "FFFFDC47"},
            "last": {"rgb": "FFFFEB9C"},
            "high": {"rgb": "FF60D276"},
            "low": {"rgb": "FFFF5367"},
        },  # 33
        {
            "series": {"rgb": "FF00B050"},
            "negative": {"rgb": "FFFF0000"},
            "markers": {"rgb": "FF0070C0"},
            "first": {"rgb": "FFFFC000"},
            "last": {"rgb": "FFFFC000"},
            "high": {"rgb": "FF00B050"},
            "low": {"rgb": "FFFF0000"},
        },  # 34
        {
            "series": {"theme": "3"},
            "negative": {"theme": "9"},
            "markers": {"theme": "8"},
            "first": {"theme": "4"},
            "last": {"theme": "5"},
            "high": {"theme": "6"},
            "low": {"theme": "7"},
        },  # 35
        {
            "series": {"theme": "1"},
            "negative": {"theme": "9"},
            "markers": {"theme": "8"},
            "first": {"theme": "4"},
            "last": {"theme": "5"},
            "high": {"theme": "6"},
            "low": {"theme": "7"},
        },  # 36
    ]

    return styles[style_id]


def _supported_datetime(
    dt: Union[datetime.datetime, datetime.time, datetime.date],
) -> bool:
    # Determine is an argument is a supported datetime object.
    return isinstance(
        dt, (datetime.datetime, datetime.date, datetime.time, datetime.timedelta)
    )


def _remove_datetime_timezone(
    dt_obj: datetime.datetime, remove_timezone: bool
) -> datetime.datetime:
    # Excel doesn't support timezones in datetimes/times so we remove the
    # tzinfo from the object if the user has specified that option in the
    # constructor.
    if remove_timezone:
        dt_obj = dt_obj.replace(tzinfo=None)
    else:
        if dt_obj.tzinfo:
            raise TypeError(
                "Excel doesn't support timezones in datetimes. "
                "Set the tzinfo in the datetime/time object to None or "
                "use the 'remove_timezone' Workbook() option"
            )

    return dt_obj


def _datetime_to_excel_datetime(
    dt_obj: Union[datetime.time, datetime.datetime, datetime.timedelta, datetime.date],
    date_1904: bool,
    remove_timezone: bool,
) -> float:
    # Convert a datetime object to an Excel serial date and time. The integer
    # part of the number stores the number of days since the epoch and the
    # fractional part stores the percentage of the day.
    date_type = dt_obj
    is_timedelta = False

    if date_1904:
        # Excel for Mac date epoch.
        epoch = datetime.datetime(1904, 1, 1)
    else:
        # Default Excel epoch.
        epoch = datetime.datetime(1899, 12, 31)

    # We handle datetime .datetime, .date and .time objects but convert
    # them to datetime.datetime objects and process them in the same way.
    if isinstance(dt_obj, datetime.datetime):
        dt_obj = _remove_datetime_timezone(dt_obj, remove_timezone)
        delta = dt_obj - epoch
    elif isinstance(dt_obj, datetime.date):
        dt_obj = datetime.datetime.fromordinal(dt_obj.toordinal())
        delta = dt_obj - epoch
    elif isinstance(dt_obj, datetime.time):
        dt_obj = datetime.datetime.combine(epoch, dt_obj)
        dt_obj = _remove_datetime_timezone(dt_obj, remove_timezone)
        delta = dt_obj - epoch
    elif isinstance(dt_obj, datetime.timedelta):
        is_timedelta = True
        delta = dt_obj
    else:
        raise TypeError("Unknown or unsupported datetime type")

    # Convert a Python datetime.datetime value to an Excel date number.
    excel_time = delta.days + (
        float(delta.seconds) + float(delta.microseconds) / 1e6
    ) / (60 * 60 * 24)

    # The following is a workaround for the fact that in Excel a time only
    # value is represented as 1899-12-31+time whereas in datetime.datetime()
    # it is 1900-1-1+time so we need to subtract the 1 day difference.
    if (
        isinstance(date_type, datetime.datetime)
        and not isinstance(dt_obj, datetime.timedelta)
        and dt_obj.isocalendar()
        == (
            1900,
            1,
            1,
        )
    ):
        excel_time -= 1

    # Account for Excel erroneously treating 1900 as a leap year.
    if not date_1904 and not is_timedelta and excel_time > 59:
        excel_time += 1

    return excel_time


def _preserve_whitespace(string: str) -> Optional[re.Match]:
    # Check if a string has leading or trailing whitespace that requires a
    # "preserve" attribute.
    return RE_LEADING_WHITESPACE.search(string) or RE_TRAILING_WHITESPACE.search(string)


def _get_image_properties(
    filename: str, image_data: Optional[BytesIO]
) -> Tuple[str, str, float, float, float, float, str]:
    # Extract dimension information from the image file.
    height = 0.0
    width = 0.0
    x_dpi = 96.0
    y_dpi = 96.0

    if not image_data:
        # Open the image file and read in the data.
        with open(filename, "rb") as fh:
            data = fh.read()
    else:
        # Read the image data from the user supplied byte stream.
        data = image_data.getvalue()

    digest = hashlib.sha256(data).hexdigest()

    # Get the image filename without the path.
    image_name = os.path.basename(filename)

    # Look for some common image file markers.
    marker1 = unpack("3s", data[1:4])[0]
    marker2 = unpack(">H", data[:2])[0]
    marker3 = unpack("2s", data[:2])[0]
    marker4 = unpack("<L", data[:4])[0]
    marker5 = (unpack("4s", data[40:44]))[0]
    marker6 = unpack("4s", data[:4])[0]

    png_marker = b"PNG"
    bmp_marker = b"BM"
    emf_marker = b" EMF"
    gif_marker = b"GIF8"

    if marker1 == png_marker:
        (image_type, width, height, x_dpi, y_dpi) = _process_png(data)

    elif marker2 == 0xFFD8:
        (image_type, width, height, x_dpi, y_dpi) = _process_jpg(data)

    elif marker3 == bmp_marker:
        (image_type, width, height) = _process_bmp(data)

    elif marker4 == 0x9AC6CDD7:
        (image_type, width, height, x_dpi, y_dpi) = _process_wmf(data)

    elif marker4 == 1 and marker5 == emf_marker:
        (image_type, width, height, x_dpi, y_dpi) = _process_emf(data)

    elif marker6 == gif_marker:
        (image_type, width, height, x_dpi, y_dpi) = _process_gif(data)

    else:
        raise UnsupportedImageFormat(
            f"{filename}: Unknown or unsupported image file format."
        )

    # Check that we found the required data.
    if not height or not width:
        raise UndefinedImageSize(f"{filename}: no size data found in image file.")

    # Set a default dpi for images with 0 dpi.
    if x_dpi == 0:
        x_dpi = 96
    if y_dpi == 0:
        y_dpi = 96

    return image_name, image_type, width, height, x_dpi, y_dpi, digest


def _process_png(
    data: bytes,
) -> Tuple[str, float, float, float, float]:
    # Extract width and height information from a PNG file.
    offset = 8
    data_length = len(data)
    end_marker = False
    width = 0.0
    height = 0.0
    x_dpi = 96.0
    y_dpi = 96.0

    # Search through the image data to read the height and width in the
    # IHDR element. Also read the DPI in the pHYs element.
    while not end_marker and offset < data_length:
        length = unpack(">I", data[offset + 0 : offset + 4])[0]
        marker = unpack("4s", data[offset + 4 : offset + 8])[0]

        # Read the image dimensions.
        if marker == b"IHDR":
            width = unpack(">I", data[offset + 8 : offset + 12])[0]
            height = unpack(">I", data[offset + 12 : offset + 16])[0]

        # Read the image DPI.
        if marker == b"pHYs":
            x_density = unpack(">I", data[offset + 8 : offset + 12])[0]
            y_density = unpack(">I", data[offset + 12 : offset + 16])[0]
            units = unpack("b", data[offset + 16 : offset + 17])[0]

            if units == 1 and x_density > 0 and y_density > 0:
                x_dpi = x_density * 0.0254
                y_dpi = y_density * 0.0254

        if marker == b"IEND":
            end_marker = True
            continue

        offset = offset + length + 12

    return "png", width, height, x_dpi, y_dpi


def _process_jpg(data: bytes) -> Tuple[str, float, float, float, float]:
    # Extract width and height information from a JPEG file.
    offset = 2
    data_length = len(data)
    end_marker = False
    width = 0.0
    height = 0.0
    x_dpi = 96.0
    y_dpi = 96.0

    # Search through the image data to read the JPEG markers.
    while not end_marker and offset < data_length:
        marker = unpack(">H", data[offset + 0 : offset + 2])[0]
        length = unpack(">H", data[offset + 2 : offset + 4])[0]

        # Read the height and width in the 0xFFCn elements (except C4, C8
        # and CC which aren't SOF markers).
        if (
            (marker & 0xFFF0) == 0xFFC0
            and marker != 0xFFC4
            and marker != 0xFFC8
            and marker != 0xFFCC
        ):
            height = unpack(">H", data[offset + 5 : offset + 7])[0]
            width = unpack(">H", data[offset + 7 : offset + 9])[0]

        # Read the DPI in the 0xFFE0 element.
        if marker == 0xFFE0:
            units = unpack("b", data[offset + 11 : offset + 12])[0]
            x_density = unpack(">H", data[offset + 12 : offset + 14])[0]
            y_density = unpack(">H", data[offset + 14 : offset + 16])[0]

            if units == 1:
                x_dpi = x_density
                y_dpi = y_density

            if units == 2:
                x_dpi = x_density * 2.54
                y_dpi = y_density * 2.54

            # Workaround for incorrect dpi.
            if x_dpi == 1:
                x_dpi = 96
            if y_dpi == 1:
                y_dpi = 96

        if marker == 0xFFDA:
            end_marker = True
            continue

        offset = offset + length + 2

    return "jpeg", width, height, x_dpi, y_dpi


def _process_gif(data: bytes) -> Tuple[str, float, float, float, float]:
    # Extract width and height information from a GIF file.
    x_dpi = 96.0
    y_dpi = 96.0

    width = unpack("<h", data[6:8])[0]
    height = unpack("<h", data[8:10])[0]

    return "gif", width, height, x_dpi, y_dpi


def _process_bmp(data: bytes) -> Tuple[str, float, float]:
    # Extract width and height information from a BMP file.
    width = unpack("<L", data[18:22])[0]
    height = unpack("<L", data[22:26])[0]
    return "bmp", width, height


def _process_wmf(data: bytes) -> Tuple[str, float, float, float, float]:
    # Extract width and height information from a WMF file.
    x_dpi = 96.0
    y_dpi = 96.0

    # Read the bounding box, measured in logical units.
    x1 = unpack("<h", data[6:8])[0]
    y1 = unpack("<h", data[8:10])[0]
    x2 = unpack("<h", data[10:12])[0]
    y2 = unpack("<h", data[12:14])[0]

    # Read the number of logical units per inch. Used to scale the image.
    inch = unpack("<H", data[14:16])[0]

    # Convert to rendered height and width.
    width = float((x2 - x1) * x_dpi) / inch
    height = float((y2 - y1) * y_dpi) / inch

    return "wmf", width, height, x_dpi, y_dpi


def _process_emf(data: bytes) -> Tuple[str, float, float, float, float]:
    # Extract width and height information from a EMF file.

    # Read the bounding box, measured in logical units.
    bound_x1 = unpack("<l", data[8:12])[0]
    bound_y1 = unpack("<l", data[12:16])[0]
    bound_x2 = unpack("<l", data[16:20])[0]
    bound_y2 = unpack("<l", data[20:24])[0]

    # Convert the bounds to width and height.
    width = bound_x2 - bound_x1
    height = bound_y2 - bound_y1

    # Read the rectangular frame in units of 0.01mm.
    frame_x1 = unpack("<l", data[24:28])[0]
    frame_y1 = unpack("<l", data[28:32])[0]
    frame_x2 = unpack("<l", data[32:36])[0]
    frame_y2 = unpack("<l", data[36:40])[0]

    # Convert the frame bounds to mm width and height.
    width_mm = 0.01 * (frame_x2 - frame_x1)
    height_mm = 0.01 * (frame_y2 - frame_y1)

    # Get the dpi based on the logical size.
    x_dpi = width * 25.4 / width_mm
    y_dpi = height * 25.4 / height_mm

    # This is to match Excel's calculation. It is probably to account for
    # the fact that the bounding box is inclusive-inclusive. Or a bug.
    width += 1
    height += 1

    return "emf", width, height, x_dpi, y_dpi
