.. _format:

The Format Class
================

This section describes the methods and properties that are available
for formatting cells in Excel. The properties of a cell that can be
formatted include: fonts, colours, patterns, borders, alignment and
number formatting.




set_font_name( $fontname )
--------------------------

    Default state: Font is Calibri
    Default action: None
    Valid args: Any valid font name

Specify the font used::

    format.set_font_name('Times New Roman')

Excel can only display fonts that are installed on the system that it
is running on. Therefore it is best to use the fonts that come as
standard such as 'Calibri', 'Times New Roman' and 'Courier New'. See
also the Fonts worksheet created by formats.pl


set_font_size()
---------------

    Default state: Font size is 10
    Default action: Set font size to 1
    Valid args: Integer values from 1 to as big as your screen.

Set the font size. Excel adjusts the height of a row to accommodate
the largest font size in the row. You can also explicitly specify the
height of a row using the set_row() worksheet method.

    format = workbook.add_format()
    format.set_font_size(30)


set_font_color()
----------------

    Default state: Excels default color, usually black
    Default action: Set the default color
    Valid args: Integers from 8..63 or the following strings::
    
                        'black'
                        'blue'
                        'brown'
                        'cyan'
                        'gray'
                        'green'
                        'lime'
                        'magenta'
                        'navy'
                        'orange'
                        'pink'
                        'purple'
                        'red'
                        'silver'
                        'white'
                        'yellow'

Set the font colour. The ``set_font_color()`` method is used as follows::

    format = workbook.add_format()
    format.set_font_color('red')
    worksheet.write(0, 0, 'wheelbarrow', format)

Note: The ``set_font_color()`` method is used to set the colour of the
font in a cell. To set the colour of a cell use the ``set_bg_color()``
and ``set_pattern()`` methods.

For additional examples see the 'Named colors' and 'Standard colors'
worksheets created by formats.pl in the examples directory.

See also "COLOURS IN EXCEL".


set_bold()
----------

    Default state: bold is off
    Default action: Turn bold on
    Valid args: 0, 1

Set the bold property of the font::

    format.set_bold(); # Turn bold on


set_italic()
------------

    Default state: Italic is off
    Default action: Turn italic on
    Valid args: 0, 1

Set the italic property of the font::

    format.set_italic(); # Turn italic on


set_underline()
---------------

    Default state: Underline is off
    Default action: Turn on single underline
    Valid args: 0  = No underline::
    
                        1 = Single underline
                        2 = Double underline
                        33 = Single accounting underline
                        34 = Double accounting underline

Set the underline property of the font.

    format.set_underline(); # Single underline


set_font_strikeout()
--------------------

    Default state: Strikeout is off
    Default action: Turn strikeout on
    Valid args: 0, 1

Set the strikeout property of the font.


set_font_script()
-----------------

    Default state: Super/Subscript is off
    Default action: Turn Superscript on
    Valid args::
    
     0  = Normal
     1 = Superscript
     2 = Subscript

Set the superscript/subscript property of the font.


set_font_outline()
------------------

    Default state: Outline is off
    Default action: Turn outline on
    Valid args: 0, 1

Macintosh only.


set_font_shadow()
-----------------

    Default state: Shadow is off
    Default action: Turn shadow on
    Valid args: 0, 1

Macintosh only.


set_num_format()
----------------

    Default state: General format
    Default action: Format index 1
    Valid args: See the following table

This method is used to define the numerical format of a number in
Excel. It controls whether a number is displayed as an integer, a
floating point number, a date, a currency value or some other user
defined format.

The numerical format of a cell can be specified by using a format
string or an index to one of Excel's built-in formats::

    format1 = workbook.add_format()
    format2 = workbook.add_format()
    format1.set_num_format('d mmm yyyy'); # Format string
    format2.set_num_format(0x0f); # Format index

    worksheet.write(0, 0, 36892.521, format1); # 1 Jan 2001
    worksheet.write(0, 0, 36892.521, format2); # 1-Jan-01

Using format strings you can define very sophisticated formatting of numbers::

    format01.set_num_format('0.000')
    worksheet.write(0, 0, 3.1415926, format01); # 3.142

    format02.set_num_format('#,##0')
    worksheet.write(1, 0, 1234.56, format02); # 1,235

    format03.set_num_format('#,##0.00')
    worksheet.write(2, 0, 1234.56, format03); # 1,234.56

    format04.set_num_format('0.00')
    worksheet.write(3, 0, 49.99, format04); # 49.99

    # Note you can use other currency symbols such as the pound or yen as well.
    # Other currencies may require the use of Unicode.

    format07.set_num_format('mm/dd/yy')
    worksheet.write(6, 0, 36892.521, format07); # 01/01/01

    format08.set_num_format('mmm d yyyy')
    worksheet.write(7, 0, 36892.521, format08); # Jan 1 2001

    format09.set_num_format('d mmmm yyyy')
    worksheet.write(8, 0, 36892.521, format09); # 1 January 2001

    format10.set_num_format('dd/mm/yyyy hh:mm AM/PM')
    worksheet.write(9, 0, 36892.521, format10); # 01/01/2001 12:30 AM

    format11.set_num_format('0 "dollar and" .00 "cents"')
    worksheet.write(10, 0, 1.87, format11); # 1 dollar and .87 cents

    # Conditional numerical formatting.
    format12.set_num_format('[Green]General;[Red]-General;General')
    worksheet.write(11, 0, 123, format12); # > 0 Green
    worksheet.write(12, 0, -45, format12); # < 0 Red
    worksheet.write(13, 0, 0, format12); # = 0 Default colour

    # Zip code
    format13.set_num_format('00000')
    worksheet.write(14, 0, '01209', format13)

The number system used for dates is described in "DATES AND TIME IN EXCEL".

The colour format should have one of the following values::

    [Black] [Blue] [Cyan] [Green] [Magenta] [Red] [White] [Yellow]

Alternatively you can specify the colour based on a colour index as
follows: ``[Color n]``, where n is a standard Excel colour
index - 7. See the 'Standard colors' worksheet created by formats.pl.

For more information refer to the documentation on formatting in the
``docs`` directory of the XlsxWriter distro, the Excel on-line help or
http://office.microsoft.com/en-gb/assistance/HP051995001033.aspx.


You should ensure that the format string is valid in Excel prior to using it in WriteExcel.

Excel's built-in formats are shown in the following table::

    Index   Index   Format String
    0       0x00    General
    1       0x01    0
    2       0x02    0.00
    3       0x03    #,##0
    4       0x04    #,##0.00
    5       0x05    ($#,##0_);($#,##0)
    6       0x06    ($#,##0_);[Red]($#,##0)
    7       0x07    ($#,##0.00_);($#,##0.00)
    8       0x08    ($#,##0.00_);[Red]($#,##0.00)
    9       0x09    0%
    10      0x0a    0.00%
    11      0x0b    0.00E+00
    12      0x0c    # ?/?
    13      0x0d    # ??/??
    14      0x0e    m/d/yy
    15      0x0f    d-mmm-yy
    16      0x10    d-mmm
    17      0x11    mmm-yy
    18      0x12    h:mm AM/PM
    19      0x13    h:mm:ss AM/PM
    20      0x14    h:mm
    21      0x15    h:mm:ss
    22      0x16    m/d/yy h:mm
    ..      ....    ...........
    37      0x25    (#,##0_);(#,##0)
    38      0x26    (#,##0_);[Red](#,##0)
    39      0x27    (#,##0.00_);(#,##0.00)
    40      0x28    (#,##0.00_);[Red](#,##0.00)
    41      0x29    _(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)
    42      0x2a    _($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)
    43      0x2b    _(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)
    44      0x2c    _($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)
    45      0x2d    mm:ss
    46      0x2e    [h]:mm:ss
    47      0x2f    mm:ss.0
    48      0x30    ##0.0E+0
    49      0x31    @

For examples of these formatting codes see the 'Numerical formats'
worksheet created by formats.pl. See also the number_formats1.html and
the number_formats2.html documents in the ``docs`` directory of the
distro.

Note 1. Numeric formats 23 to 36 are not documented by Microsoft and
may differ in international versions.

Note 2. The dollar sign appears as the defined local currency symbol.


set_locked()
------------

    Default state: Cell locking is on
    Default action: Turn locking on
    Valid args: 0, 1

This property can be used to prevent modification of a cells
contents. Following Excel's convention, cell locking is turned on by
default. However, it only has an effect if the worksheet has been
protected, see the worksheet ``protect()`` method.

    locked = workbook.add_format()
    locked.set_locked(1); # A non-op

    unlocked = workbook.add_format()
    locked.set_locked(0)

    # Enable worksheet protection
    worksheet.protect()

    # This cell cannot be edited.
    worksheet.write('A1', '=1+2', locked)

    # This cell can be edited.
    worksheet.write('A2', '=1+2', unlocked)

Note: This offers weak protection even with a password, see the note
in relation to the ``protect()`` method.


set_hidden()
------------

    Default state: Formula hiding is off
    Default action: Turn hiding on
    Valid args: 0, 1

This property is used to hide a formula while still displaying its
result. This is generally used to hide complex calculations from end
users who are only interested in the result. It only has an effect if
the worksheet has been protected, see the worksheet ``protect()``
method.

    hidden = workbook.add_format()
    hidden.set_hidden()

    # Enable worksheet protection
    worksheet.protect()

    # The formula in this cell isn't visible
    worksheet.write('A1', '=1+2', hidden)

Note: This offers weak protection even with a password, see the note
in relation to the ``protect()`` method.



set_align()
-----------

    Default state: Alignment is off
    Default action: Left alignment
    Valid args::
    
                       'left' Horizontal
                        'center'
                        'right'
                        'fill'
                        'justify'
                        'center_across'

                        'top' Vertical
                        'vcenter'
                        'bottom'
                        'vjustify'

This method is used to set the horizontal and vertical text alignment
within a cell. Vertical and horizontal alignments can be combined. The
method is used as follows::

    format = workbook.add_format()
    format.set_align('center')
    format.set_align('vcenter')
    worksheet.set_row(0, 30)
    worksheet.write(0, 0, 'X', format)

Text can be aligned across two or more adjacent cells using the
``center_across`` property. However, for genuine merged cells it is
better to use the ``merge_range()`` worksheet method.

The ``vjustify`` (vertical justify) option can be used to provide
automatic text wrapping in a cell. The height of the cell will be
adjusted to accommodate the wrapped text. To specify where the text
wraps use the ``set_text_wrap()`` method.

For further examples see the 'Alignment' worksheet created by
formats.pl.


set_center_across()
-------------------

    Default state: Center across selection is off
    Default action: Turn center across on
    Valid args: 1

Text can be aligned across two or more adjacent cells using the
``set_center_across()`` method. This is an alias for the
``set_align('center_across')`` method call.

Only one cell should contain the text, the other cells should be
blank::


    format = workbook.add_format()
    format.set_center_across()

    worksheet.write(1, 1, 'Center across selection', format)
    worksheet.write_blank(1, 2, format)

See also the ``merge1.pl`` to ``merge6.pl`` programs in the
``examples`` directory and the ``merge_range()`` method.



set_text_wrap()
---------------

    Default state: Text wrap is off
    Default action: Turn text wrap on
    Valid args: 0, 1

Here is an example using the text wrap property, the escape character
``\n`` is used to indicate the end of line::


    format = workbook.add_format()
    format.set_text_wrap()
    worksheet.write(0, 0, "It's\na bum\nwrap", format)

Excel will adjust the height of the row to accommodate the wrapped
text. A similar effect can be obtained without newlines using the
``set_align('vjustify')`` method. See the ``textwrap.pl`` program in
the ``examples`` directory.



set_rotation()
--------------

    Default state: Text rotation is off
    Default action: None
    Valid args: Integers in the range -90 to 90 and 270

Set the rotation of the text in a cell. The rotation can be any angle
in the range -90 to 90 degrees.


    format = workbook.add_format()
    format.set_rotation(30)
    worksheet.write(0, 0, 'This text is rotated', format)

The angle 270 is also supported. This indicates text where the letters
run from top to bottom.



set_indent()
------------

    Default state: Text indentation is off
    Default action: Indent text 1 level
    Valid args: Positive integers

This method can be used to indent text. The argument, which should be
an integer, is taken as the level of indentation::


    format = workbook.add_format()
    format.set_indent(2)
    worksheet.write(0, 0, 'This text is indented', format)

Indentation is a horizontal alignment property. It will override any
other horizontal properties but it can be used in conjunction with
vertical properties.



set_shrink()
------------

    Default state: Text shrinking is off
    Default action: Turn "shrink to fit" on
    Valid args: 1

This method can be used to shrink text so that it fits in a cell.

    format = workbook.add_format()
    format.set_shrink()
    worksheet.write(0, 0, 'Honey, I shrunk the textnot ', format)


set_text_justlast()
-------------------

    Default state: Justify last is off
    Default action: Turn justify last on
    Valid args: 0, 1

Only applies to Far Eastern versions of Excel.


set_pattern()
-------------

    Default state: Pattern is off
    Default action: Solid fill is on
    Valid args: 0 .. 18

Set the background pattern of a cell.

Examples of the available patterns are shown in the 'Patterns'
worksheet created by formats.pl. However, it is unlikely that you will
ever need anything other than Pattern 1 which is a solid fill of the
background color.



set_bg_color()
--------------

    Default state: Color is off
    Default action: Solid fill.
    Valid args: See set_font_color()

The ``set_bg_color()`` method can be used to set the background colour
of a pattern. Patterns are defined via the ``set_pattern()``
method. If a pattern hasn't been defined then a solid fill pattern is
used as the default.


Here is an example of how to set up a solid fill in a cell::

    format = workbook.add_format()

    format.set_pattern(); # This is optional when using a solid fill

    format.set_bg_color('green')
    worksheet.write('A1', 'Ray', format)

For further examples see the 'Patterns' worksheet created by formats.pl.


set_fg_color()
--------------

    Default state: Color is off
    Default action: Solid fill.
    Valid args: See set_font_color()

The ``set_fg_color()`` method can be used to set the foreground colour
of a pattern.


For further examples see the 'Patterns' worksheet created by
formats.pl.



set_border()
------------

    Also applies to: set_bottom()
                        set_top()
                        set_left()
                        set_right()

    Default state: Border is off
    Default action: Set border type 1
    Valid args: 0-13, See below.

A cell border is comprised of a border on the bottom, top, left and
right. These can be set to the same value using ``set_border()`` or
individually using the relevant method calls shown above.

The following shows the border styles sorted by XlsxWriter index
number::

    Index   Name            Weight   Style
    =====   =============   ======   ===========
    0       None            0
    1       Continuous      1        -----------
    2       Continuous      2        -----------
    3       Dash            1        - - - - - -
    4       Dot             1        . . . . . .
    5       Continuous      3        -----------
    6       Double          3        ===========
    7       Continuous      0        -----------
    8       Dash            2        - - - - - -
    9       Dash Dot        1        - . - . - .
    10      Dash Dot        2        - . - . - .
    11      Dash Dot Dot    1        - . . - . .
    12      Dash Dot Dot    2        - . . - . .
    13      SlantDash Dot   2        / - . / - .


The following shows the borders sorted by style::

    Name            Weight   Style         Index
    =============   ======   ===========   =====
    Continuous      0        -----------   7
    Continuous      1        -----------   1
    Continuous      2        -----------   2
    Continuous      3        -----------   5
    Dash            1        - - - - - -   3
    Dash            2        - - - - - -   8
    Dash Dot        1        - . - . - .   9
    Dash Dot        2        - . - . - .   10
    Dash Dot Dot    1        - . . - . .   11
    Dash Dot Dot    2        - . . - . .   12
    Dot             1        . . . . . .   4
    Double          3        ===========   6
    None            0                      0
    SlantDash Dot   2        / - . / - .   13


The following shows the borders in the order shown in the Excel
Dialog::


    Index   Style             Index   Style
    =====   =====             =====   =====
    0       None              12      - . . - . .
    7       -----------       13      / - . / - .
    4       . . . . . .       10      - . - . - .
    11      - . . - . .       8       - - - - - -
    9       - . - . - .       2       -----------
    3       - - - - - -       5       -----------
    1       -----------       6       ===========


set_border_color()
------------------

    Also applies to: set_bottom_color()
                        set_top_color()
                        set_left_color()
                        set_right_color()

    Default state: Color is off
    Default action: Undefined
    Valid args: See set_font_color()

Set the colour of the cell borders. A cell border is comprised of a
border on the bottom, top, left and right. These can be set to the
same colour using ``set_border_color()`` or individually using the
relevant method calls shown above. Examples of the border styles and
colours are shown in the 'Borders' worksheet created by formats.pl.


