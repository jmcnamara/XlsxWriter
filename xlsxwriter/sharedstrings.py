###############################################################################
#
# SharedStrings - A class for writing the Excel XLSX sharedStrings file.
#
# Copyright 2013, John McNamara, jmcnamara@cpan.org
#

import re
import xmlwriter


class SharedStrings(xmlwriter.XMLwriter):
    """
    A class for writing the Excel XLSX sharedStrings file.

    """

    ###########################################################################
    #
    # Public API.
    #
    ###########################################################################

    def __init__(self):
        """
        Construtor.

        """

        super(SharedStrings, self).__init__()

        self.string_table = None

    ###########################################################################
    #
    # Private API.
    #
    ###########################################################################

    def _assemble_xml_file(self):
        # Assemble and write the XML file.

        # Write the XML declaration.
        self._xml_declaration()

        # Write the sst element.
        self._write_sst()

        # Write the sst strings.
        self._write_sst_strings()

        # Close the sst tag.
        self._xml_end_tag('sst')

        # Close the file.
        self._xml_close()

    ###########################################################################
    #
    # XML methods.
    #
    ###########################################################################

    def _write_sst(self):
        # Write the <sst> element.
        xmlns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

        attributes = [
            ('xmlns', xmlns),
            ('count', self.string_table.count),
            ('uniqueCount', self.string_table.unique_count),
        ]

        self._xml_start_tag('sst', attributes)

    def _write_sst_strings(self):
        # Write the sst string elements.

        for string in (self.string_table._get_strings()):
            self._write_si(string)

    def _write_si(self, string):
        # Write the <si> element.
        attributes = []

        # TODO: Fix control char encoding when unit test is ported.
        # Excel escapes control characters with _xHHHH_ and also escapes any
        # literal strings of that type by encoding the leading underscore.
        # So "\0" -> _x0000_ and "_x0000_" -> _x005F_x0000_.
        # The following substitutions deal with those cases.

        # Escape the escape.
        # string =~ s/(_x[0-9a-fA-F]{4}_)/_x005F1/g

        # Convert control character to the _xHHHH_ escape.
        # string =~ s/([\x00-\x08\x0B-\x1F])/sprintf "_x04X_", ord(1)/eg

        # Add attribute to preserve leading or trailing whitespace.
        if re.search('^\s', string) or re.search('\s$', string):
            attributes.append(('xml:space', 'preserve'))

        # Write any rich strings without further tags.
        if re.search('^<r>', string) and re.search('</r>$', string):
            # Prevent utf8 strings from getting double encoded.
            # string = decode_utf8(string)

            self._xml_rich_si_element(string)
        else:
            self._xml_si_element(string, attributes)


# A metadata class to store Excel strings between worksheets.
class SharedStringTable(object):
    """
    A class to track Excel shared strings between worksheets.

    """

    def __init__(self):
        self.count = 0
        self.unique_count = 0
        self.string_table = {}

    def _get_shared_string_index(self, string):
        """" Get the index of the string in the Shared String table. """
        if string not in self.string_table:
            # String isn't already stored in the table so add it.
            index = self.unique_count
            self.string_table[string] = index
            self.count += 1
            self.unique_count += 1
            return index
        else:
            # String exists in the table.
            index = self.string_table[string]
            self.count += 1
            return index

    def _get_strings(self):
        return sorted(self.string_table, key=self.string_table.__getitem__)
