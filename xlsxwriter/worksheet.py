###############################################################################
#
# Worksheet - A class for writing Excel Worksheets.
#
# Copyright 2013, John McNamara, jmcnamara@cpan.org
#

import xmlwriter


class Worksheet(xmlwriter.XMLwriter):
    """
    A class for writing Excel Worksheets.

    """

    ###########################################################################
    #
    # Public API.
    #

    ###########################################################################
    #
    # Private API.
    #

    def _assemble_xml_file(self):
        # Assemble and write the XML file.

        self._xml_declaration()

    ###########################################################################
    #
    # XML methods.
    #
