###############################################################################
#
# Worksheet - A class for writing Excel Worksheets.
#
# Copyright 2013, John McNamara, jmcnamara@cpan.org
#


def xl_rowcol_to_cell(row, col, row_abs=0, col_abs=0):
    """
    TODO

    """
    row = row + 1  # Change to 1-index.
    row_abs = '$' if row_abs else ''
    col_abs = '$' if col_abs else ''

    col_str = xl_col_to_name(col, col_abs)

    return col_str + row_abs + str(row)


def xl_col_to_name(col_num, col_abs=0):
    """
    TODO

    """
    col_num = col_num + 1  # Change to 1-index.
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
