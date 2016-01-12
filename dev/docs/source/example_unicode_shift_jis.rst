.. _ex_unicode_shift_jis:

Example: Unicode - Shift JIS
============================

This program is an example of reading in data from a Shift JIS encoded text
file and converting it to a worksheet.

The main trick is to ensure that the data read in is converted to UTF-8
within the Python program. The XlsxWriter module will then take care of
writing the encoding to the Excel file.

The encoding of the input data shouldn't matter once it can be converted
to UTF-8 via the :mod:`codecs` module.

.. image:: _images/unicode_shift_jis.png

.. literalinclude:: ../../../examples/unicode_shift_jis.py

