.. _ex_merge_rich:

Example: Merging Cells with a Rich String
=========================================

This program is an example of merging cells that contain a rich string.

Using the standard XlsxWriter API we can only write simple types to merged
ranges so we first write a blank string to the merged range. We then overwrite
the first merged cell with a rich string.

Note that we must also pass the cell format used in the merged cells format at
the end

See the :func:`merge_range` and :func:`write_rich_string` methods for more
details.

.. image:: _images/merge_rich.png

.. literalinclude:: ../../../examples/merge_rich_string.py

