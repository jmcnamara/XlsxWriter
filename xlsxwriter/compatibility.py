###############################################################################
#
# Python 2/3 compatibility functions for XlsxWriter.
#
# Copyright (c), 2013, John McNamara, jmcnamara@cpan.org
#

try:
    # For compatibility between Python 2 and 3.
    from StringIO import StringIO
except ImportError:
    from io import StringIO

try:
    # For Python 2.6+.
    from fractions import Fraction
except ImportError:
    Fraction = float

try:
    # For Python 2.6+.
    from collections import defaultdict
    from collections import namedtuple
except ImportError:
    # For Python 2.5 support.
    from .compat_collections import defaultdict
    from .compat_collections import namedtuple
