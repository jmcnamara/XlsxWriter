.. SPDX-License-Identifier: BSD-2-Clause
   Copyright (c) 2013-2025, John McNamara, jmcnamara@cpan.org

.. _ex_embedded_images:

Example: Embedding images into a worksheet
==========================================

This program is an example of embedding images into a worksheet. The image will
scale automatically to fit the cell.

This is the equivalent of Excel's menu option to insert an image using the
option to "Place in Cell" which is only available in Excel 365 versions from
2023 onwards. For older versions of Excel a ``#VALUE!`` error is displayed.

See the
:func:`embed_image` method for more details.

.. image:: _images/embedded_images.png

.. literalinclude:: ../../../examples/embedded_images.py

