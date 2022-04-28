.. SPDX-License-Identifier: BSD-2-Clause
   Copyright 2013-2022, John McNamara, jmcnamara@cpan.org

.. _ex_lambda:

Example: Excel 365 LAMBDA() function
====================================

This program is an example of using the new Excel ``LAMBDA()`` function. It
demonstrates how to create a lambda function in Excel and also how to assign a
name to it so that it can be called as a user defined function. This
particular example converts from Fahrenheit to Celsius.

Note, this function is only currently available if you
are subscribed to the Microsoft Office Beta Channel program.  See the
:ref:`formula_lambda` section of the documentation for more details.

.. image:: _images/lambda01.png

.. literalinclude:: ../../../examples/lambda.py

