.. SPDX-License-Identifier: BSD-2-Clause
   Copyright 2013-2023, John McNamara, jmcnamara@cpan.org

.. _ex_protection:

Example: Enabling Cell protection in Worksheets
===============================================

This program is an example cell locking and formula hiding in an Excel
worksheet using the :func:`protect` worksheet method and the Format
:func:`set_locked` property.

Note, that Excel's behavior is that all cells are locked once you set the
default protection. Therefore you need to explicitly unlock cells rather than
explicitly lock them.

.. image:: _images/worksheet_protection.png

.. literalinclude:: ../../../examples/worksheet_protection.py

