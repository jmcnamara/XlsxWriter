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

    