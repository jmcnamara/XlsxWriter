.. _ex_demo:

Example: Simple Feature Demonstration
=====================================

This program is an example of writing some of the features of the
XlsxWriter module.

.. image:: _static/demo.png

Code:

.. literalinclude:: ../../../examples/demo.py

Notes:

* This example includes the use of cell formatting via the :ref:`format`.
* Strings and numbers can be written with the same worksheet :func:`write`
  method.
* Data can be written to cells using Row-Column notation or 'A1' style
  notation, see :ref:`cell_notation`.