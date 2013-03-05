###############################################################################
#
# Worksheet - A class for writing Excel Worksheets.
#
# Copyright 2013, John McNamara, jmcnamara@cpan.org
#
import re


range_parts = re.compile(r'(\$?)([A-Z]{1,3})(\$?)(\d+)')


def xl_rowcol_to_cell(row, col, row_abs=0, col_abs=0):
    """
    TODO. Add Utility.py docs.

    """
    row += 1  # Change to 1-index.
    row_abs = '$' if row_abs else ''
    col_abs = '$' if col_abs else ''

    col_str = xl_col_to_name(col, col_abs)

    return col_str + row_abs + str(row)


def xl_col_to_name(col_num, col_abs=0):
    """
    TODO. Add Utility.py docs.

    """
    col_num += 1  # Change to 1-index.
    col_str = ''
    col_abs = '$' if col_abs else ''

    while col_num:

        # Set remainder from 1 .. 26
        remainder = col_num % 26

        if remainder == 0:
            remainder = 26

        # Convert the remainder to a character.
        col_letter = chr(ord('A') + remainder - 1)

        # Accumulate the column letters, right to left.
        col_str = col_letter + col_str

        # Get the next order of magnitude.
        col_num = int((col_num - 1) / 26)

    return col_abs + col_str


def xl_cell_to_rowcol(cell_str):
    """
    TODO. Add Utility.py docs.

    """
    if not cell_str:
        return (0, 0)

    match = range_parts.match(cell_str)
    col_str = match.group(2)
    row_str = match.group(4)

    # Convert base26 column string to number.
    expn = 0
    col = 0
    for char in reversed(col_str):
        col += (ord(char) - ord('A') + 1) * (26 ** expn)
        expn += 1

    # Convert 1-index to zero-index
    row = int(row_str) - 1
    col -= 1

    return row, col


def xl_cell_to_rowcol_abs(cell_str):
    """
    TODO. Add Utility.py docs.

    """
    if not cell_str:
        return (0, 0, 0, 0)

    match = range_parts.match(cell_str)

    col_abs = match.group(1)
    col_str = match.group(2)
    row_abs = match.group(3)
    row_str = match.group(4)

    if col_abs:
        col_abs = 1
    else:
        col_abs = 0

    if row_abs:
        row_abs = 1
    else:
        row_abs = 0

    # Convert base26 column string to number.
    expn = 0
    col = 0
    for char in reversed(col_str):
        col += (ord(char) - ord('A') + 1) * (26 ** expn)
        expn += 1

    # Convert 1-index to zero-index
    row = int(row_str) - 1
    col -= 1

    return row, col, row_abs, col_abs


def xl_range(first_row, first_col, last_row, last_col):
    """
    TODO. Add Utility.py docs.

    """
    range1 = xl_rowcol_to_cell(first_row, first_col)
    range2 = xl_rowcol_to_cell(last_row, last_col)

    return range1 + ':' + range2
