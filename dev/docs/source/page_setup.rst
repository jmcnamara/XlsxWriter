.. _page_setup:

The Worksheet Class (Page Setup)
================================

Page set-up methods affect the way that a worksheet looks when it is printed.
They control features such as paper size, orientation, page headers and
margins.

These methods are really just standard :ref:`worksheet <worksheet>` methods.
They are documented separately for the sake of clarity.


worksheet.set_landscape()
-------------------------

.. py:function:: set_landscape()

   Set the page orientation as landscape.

This method is used to set the orientation of a worksheet's printed page to
landscape::

    worksheet.set_landscape()


worksheet.set_portrait()
------------------------

.. py:function:: set_portrait()

   Set the page orientation as portrait.

This method is used to set the orientation of a worksheet's printed page to
portrait. The default worksheet orientation is portrait, so you won't
generally need to call this method::

    worksheet.set_portrait()


worksheet.set_page_view()
-------------------------

.. py:function:: set_page_view()

   Set the page view mode.

This method is used to display the worksheet in "Page View/Layout" mode::

    worksheet.set_page_view()


worksheet.set_paper()
---------------------

.. py:function:: set_paper(index)

   Set the paper type.

   :param int index: The Excel paper format index.

This method is used to set the paper format for the printed output of a
worksheet. The following paper styles are available:

    =====   ====================    ================
    Index   Paper format            Paper size
    =====   ====================    ================
      0     Printer default
      1     Letter                  8 1/2 x 11 in
      2     Letter Small            8 1/2 x 11 in
      3     Tabloid                 11 x 17 in
      4     Ledger                  17 x 11 in
      5     Legal                   8 1/2 x 14 in
      6     Statement               5 1/2 x 8 1/2 in
      7     Executive               7 1/4 x 10 1/2 in
      8     A3                      297 x 420 mm
      9     A4                      210 x 297 mm
     10     A4 Small                210 x 297 mm
     11     A5                      148 x 210 mm
     12     B4                      250 x 354 mm
     13     B5                      182 x 257 mm
     14     Folio                   8 1/2 x 13 in
     15     Quarto                  215 x 275 mm
     16                             10x14 in
     17                             11x17 in
     18     Note                    8 1/2 x 11 in
     19     Envelope  9             3 7/8 x 8 7/8
     20     Envelope 10             4 1/8 x 9 1/2
     21     Envelope 11             4 1/2 x 10 3/8
     22     Envelope 12             4 3/4 x 11
     23     Envelope 14             5 x 11 1/2
     24     C size sheet
     25     D size sheet
     26     E size sheet
     27     Envelope DL             110 x 220 mm
     28     Envelope C3             324 x 458 mm
     29     Envelope C4             229 x 324 mm
     30     Envelope C5             162 x 229 mm
     31     Envelope C6             114 x 162 mm
     32     Envelope C65            114 x 229 mm
     33     Envelope B4             250 x 353 mm
     34     Envelope B5             176 x 250 mm
     35     Envelope B6             176 x 125 mm
     36     Envelope                110 x 230 mm
     37     Monarch                 3.875 x 7.5 in
     38     Envelope                3 5/8 x 6 1/2 in
     39     Fanfold                 14 7/8 x 11 in
     40     German Std Fanfold      8 1/2 x 12 in
     41     German Legal Fanfold    8 1/2 x 13 in
    =====   ====================    ================


Note, it is likely that not all of these paper types will be available to the
end user since it will depend on the paper formats that the user's printer
supports. Therefore, it is best to stick to standard paper types::

    worksheet.set_paper(1)  # US Letter
    worksheet.set_paper(9)  # A4

If you do not specify a paper type the worksheet will print using the printer's
default paper style.


center_horizontally()
---------------------

.. py:function:: center_horizontally()

   Center the printed page horizontally.

Center the worksheet data horizontally between the margins on the printed page::

    worksheet.center_horizontally()


center_vertically()
-------------------

.. py:function:: center_vertically()

   Center the printed page vertically.

Center the worksheet data vertically between the margins on the printed page::

    worksheet.center_vertically()

worksheet.set_margins()
-----------------------

.. py:function:: set_margins([left=0.7,] right=0.7,] top=0.75,] bottom=0.75]]])

   Set the worksheet margins for the printed page.

   :param float left:   Left margin in inches. Default 0.7.
   :param float right:  Right margin in inches. Default 0.7.
   :param float top:    Top margin in inches. Default 0.75.
   :param float bottom: Bottom margin in inches. Default 0.75.


The ``set_margins()`` method is used to set the margins of the worksheet when
it is printed. The units are in inches. All parameters are optional and have
default values corresponding to the default Excel values.


set_header()
------------

.. py:function:: set_header([header='',] margin=0.3]])

   Set the printed page header caption and optional margin.

   :param string header: Header string with Excel control characters.
   :param float margin:  Header margin in inches. Default 0.3.

Headers and footers are generated using a string which is a combination
of plain text and control characters.

The available control character are:

    ==================  ==================  ================================
    Control             Category            Description
    ==================  ==================  ================================
    &L                  Justification       Left
    &C                                      Center
    &R                                      Right

    &P                  Information         Page number
    &N                                      Total number of pages
    &D                                      Date
    &T                                      Time
    &F                                      File name
    &A                                      Worksheet name
    &Z                                      Workbook path

    &fontsize           Font                Font size
    &"font,style"                           Font name and style
    &U                                      Single underline
    &E                                      Double underline
    &S                                      Strikethrough
    &X                                      Superscript
    &Y                                      Subscript

    &&                  Miscellaneous       Literal ampersand &
    ==================  ==================  ================================

Text in headers and footers can be justified (aligned) to the left, center and
right by prefixing the text with the control characters ``&L``, ``&C`` and
``&R``.

For example (with ASCII art representation of the results)::

    worksheet.set_header('&LHello')

        ---------------------------------------------------------------
       |                                                               |
       | Hello                                                         |
       |                                                               |


    $worksheet->set_header('&CHello');

        ---------------------------------------------------------------
       |                                                               |
       |                          Hello                                |
       |                                                               |


    $worksheet->set_header('&RHello');

        ---------------------------------------------------------------
       |                                                               |
       |                                                         Hello |
       |                                                               |


For simple text, if you do not specify any justification the text will be
centred. However, you must prefix the text with ``&C`` if you specify a font
name or any other formatting::

    worksheet.set_header('Hello')

        ---------------------------------------------------------------
       |                                                               |
       |                          Hello                                |
       |                                                               |

You can have text in each of the justification regions::

    worksheet.set_header('&LCiao&CBello&RCielo')

        ---------------------------------------------------------------
       |                                                               |
       | Ciao                     Bello                          Cielo |
       |                                                               |


The information control characters act as variables that Excel will update as
the workbook or worksheet changes. Times and dates are in the users default
format::

    worksheet.set_header('&CPage &P of &N')

        ---------------------------------------------------------------
       |                                                               |
       |                        Page 1 of 6                            |
       |                                                               |

    worksheet.set_header('&CUpdated at &T')

        ---------------------------------------------------------------
       |                                                               |
       |                    Updated at 12:30 PM                        |
       |                                                               |

You can specify the font size of a section of the text by prefixing it with the
control character ``&n`` where ``n`` is the font size::

    worksheet1.set_header('&C&30Hello Big')
    worksheet2.set_header('&C&10Hello Small')

You can specify the font of a section of the text by prefixing it with the
control sequence ``&"font,style"`` where ``fontname`` is a font name such as
"Courier New" or "Times New Roman" and ``style`` is one of the standard
Windows font descriptions: "Regular", "Italic", "Bold" or "Bold Italic"::

    worksheet1.set_header('&C&"Courier New,Italic"Hello')
    worksheet2.set_header('&C&"Courier New,Bold Italic"Hello')
    worksheet3.set_header('&C&"Times New Roman,Regular"Hello')

It is possible to combine all of these features together to create
sophisticated headers and footers. As an aid to setting up complicated headers
and footers you can record a page set-up as a macro in Excel and look at the
format strings that VBA produces. Remember however that VBA uses two double
quotes ``""`` to indicate a single double quote. For the last example above
the equivalent VBA code looks like this::

    .LeftHeader = ""
    .CenterHeader = "&""Times New Roman,Regular""Hello"
    .RightHeader = ""

To include a single literal ampersand ``&`` in a header or footer you should
use a double ampersand ``&&``::

    worksheet1.set_header('&CCuriouser and Curiouser - Attorneys at Law')

As stated above the margin parameter is optional. As with the other margins the
value should be in inches. The default header and footer margin is 0.3 inch.
The header and footer margin size can be set as follows::

    worksheet.set_header('&CHello', 0.75)

The header and footer margins are independent of the top and bottom margins.

Note, the header or footer string must be less than 255 characters. Strings
longer than this will not be written and an exception will be thrown.


set_footer()
------------

.. py:function:: set_header([footer='',] margin=0.3]])

   Set the printed page footer caption and optional margin.

   :param string footer: Footer string with Excel control characters.
   :param float margin:  Footer margin in inches. Default 0.3.

The syntax of the ``set_footer()`` method is the same as :func:`set_header`.


repeat_rows()
-------------

.. py:function:: repeat_rows(first_row[, last_row])

   Set the number of rows to repeat at the top of each printed page.

   :param int first_row: First row of repeat range.
   :param int last_row:  Last row of repeat range. Optional.

For large Excel documents it is often desirable to have the first row or rows
of the worksheet print out at the top of each page.

This can be achieved by using the ``repeat_rows()`` method. The parameters
``first_row`` and ``last_row`` are zero based. The ``last_row`` parameter is
optional if you only wish to specify one row::

    worksheet1.repeat_rows(0)     # Repeat the first row.
    worksheet2.repeat_rows(0, 1)  # Repeat the first two rows.


repeat_columns()
----------------

.. py:function:: repeat_columns(first_col[, last_col])

   Set the columns to repeat at the left hand side of each printed page.

   :param int first_col: First column of repeat range.
   :param int last_col:  Last column of repeat range. Optional.

For large Excel documents it is often desirable to have the first column or
columns of the worksheet print out at the left hand side of each page.

This can be achieved by using the ``repeat_columns()`` method. The parameters
``first_column`` and ``last_column`` are zero based. The ``last_column``
parameter is optional if you only wish to specify one column. You can also
specify the columns using A1 column notation, see :ref:`cell_notation` for
more details.::

    worksheet1.repeat_columns(0)      # Repeat the first column.
    worksheet2.repeat_columns(0, 1)   # Repeat the first two columns.
    worksheet3.repeat_columns('A:A')  # Repeat the first column.
    worksheet4.repeat_columns('A:B')  # Repeat the first two columns.


hide_gridlines()
----------------

.. py:function:: set_header([option=1])

   Set the option to hide gridlines on the screen and the printed page.

   :param int option: Hide gridline options. See below.

This method is used to hide the gridlines on the screen and printed page.
Gridlines are the lines that divide the cells on a worksheet. Screen and
printed gridlines are turned on by default in an Excel worksheet.

If you have defined your own cell borders you may wish to hide the default
gridlines::

    worksheet.hide_gridlines()

The following values of ``option`` are valid:

0. Don't hide gridlines.
1. Hide printed gridlines only.
2. Hide screen and printed gridlines.

If you don't supply an argument the default option is 1, i.e. only the
printed gridlines are hidden.


print_row_col_headers()
---------------------------

.. py:function:: print_row_col_headers()

   Set the option to print the row and column headers on the printed page.

When you print a worksheet from Excel you get the data selected in the print
area. By default the Excel row and column headers (the row numbers on the left
and the column letters at the top) aren't printed.

The ``print_row_col_headers()`` method sets the printer option to print these
headers::

    worksheet.print_row_col_headers()


worksheet.print_across()
------------------------

.. py:function:: print_across()

   Set the order in which pages are printed.

The ``print_across`` method is used to change the default print direction. This
is referred to by Excel as the sheet "page order"::

    worksheet.print_across()

The default page order is shown below for a worksheet that extends over 4
pages. The order is called "down then across"::

    [1] [3]
    [2] [4]

However, by using the ``print_across`` method the print order will be changed
to "across then down"::

    [1] [2]
    [3] [4]

    