.. SPDX-License-Identifier: BSD-2-Clause
   Copyright 2013-2024, John McNamara, jmcnamara@cpan.org

.. _ex_autofit_manually:

Example: Autofitting columns manually
=====================================

An example of simulating autofitting column widths using the
:func:`cell_autofit_width` utility function.

The following example demonstrates manually auto-fitting the the width of a
column in Excel based on the maximum string width. The worksheet :func:`autofit`
method will do this automatically but occasionally you may need to control the
maximum and minimum column widths yourself.

.. image:: _images/autofit_manually.png

.. literalinclude:: ../../../examples/autofit_manually.py

