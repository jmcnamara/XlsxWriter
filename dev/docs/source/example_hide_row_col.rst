.. _ex_hide_row_col:

Example: Hiding Rows and Columns
================================

This program is an example of how to hide rows and columns in XlsxWriter.

An individual row can be hidden using the :func:`set_row` method::

    worksheet.set_row(0, None, None, {'hidden': True})

However, in order to hide a large number of rows, for example all the rows
after row 8, we need to use an Excel optimization to hide rows without setting
each one, (of approximately 1 million rows). To do this we use the
:func:`set_default_row` method.

Columns don't require this optimization and can be hidden using
:func:`set_column`.

.. image:: _images/hide_row_col.png

.. literalinclude:: ../../../examples/hide_row_col.py

