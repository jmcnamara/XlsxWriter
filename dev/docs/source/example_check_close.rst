.. SPDX-License-Identifier: BSD-2-Clause
   Copyright 2013-2021, John McNamara, jmcnamara@cpan.org

.. _ex_check_close:

Example: Catch exception on closing
===================================

A simple program demonstrating a check for exceptions when closing the file.

We try to :func:`close()` the file in a loop so that if there is an exception,
such as if the file is open or locked, we can ask the user to close the file,
after which we can try again to overwrite it.

.. literalinclude:: ../../../examples/check_close.py
