###############################################################################
#
# Python 2/3 compatibility functions for testing XlsxWriter.
#
# Copyright (c), 2013, John McNamara, jmcnamara@cpan.org
#

try:
    from StringIO import StringIO
except ImportError:
    from io import StringIO
