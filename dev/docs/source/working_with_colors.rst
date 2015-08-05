.. _colors:

Working with Colors
===================

Throughout XlsxWriter colors are specified using a Html style ``#RRGGBB``
value. For example with a :ref:`Format <format>` object::

    cell_format.set_font_color('#FF0000')

For backward compatibility a limited number of color names are supported::

    cell_format.set_font_color('red')

The color names and corresponding ``#RRGGBB`` value are shown below:

+------------+----------------+
| Color name | RGB color code |
+============+================+
| black      | ``#000000``    |
+------------+----------------+
| blue       | ``#0000FF``    |
+------------+----------------+
| brown      | ``#800000``    |
+------------+----------------+
| cyan       | ``#00FFFF``    |
+------------+----------------+
| gray       | ``#808080``    |
+------------+----------------+
| green      | ``#008000``    |
+------------+----------------+
| lime       | ``#00FF00``    |
+------------+----------------+
| magenta    | ``#FF00FF``    |
+------------+----------------+
| navy       | ``#000080``    |
+------------+----------------+
| orange     | ``#FF6600``    |
+------------+----------------+
| pink       | ``#FF00FF``    |
+------------+----------------+
| purple     | ``#800080``    |
+------------+----------------+
| red        | ``#FF0000``    |
+------------+----------------+
| silver     | ``#C0C0C0``    |
+------------+----------------+
| white      | ``#FFFFFF``    |
+------------+----------------+
| yellow     | ``#FFFF00``    |
+------------+----------------+
