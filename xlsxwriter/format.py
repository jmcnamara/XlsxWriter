###############################################################################
#
# Format - A class for writing the Excel XLSX Worksheet file.
#
# Copyright 2013, John McNamara, jmcnamara@cpan.org
#

import xmlwriter


class Format(xmlwriter.XMLwriter):
    """
    A class for writing the Excel XLSX Format file.


    """

    ###########################################################################
    #
    # Public API.
    #
    ###########################################################################

    def __init__(self):
        """
        Constructor.

        """

        super(Format, self).__init__()

        self.xf_format_indices = None
        self.dxf_format_indices = None
        self.xf_index = None
        self.dxf_index = None

        self.num_format = 0
        self.num_format_index = 0
        self.font_index = 0
        self.has_font = 0
        self.has_dxf_font = 0

        self.font = 'Calibri'
        self.size = 11
        self.bold = 0
        self.italic = 0
        self.color = 0x0
        self.underline = 0
        self.font_strikeout = 0
        self.font_outline = 0
        self.font_shadow = 0
        self.font_script = 0
        self.font_family = 2
        self.font_charset = 0
        self.font_scheme = 'minor'
        self.font_condense = 0
        self.font_extend = 0
        self.theme = 0
        self.hyperlink = 0

        self.hidden = 0
        self.locked = 1

        self.text_h_align = 0
        self.text_wrap = 0
        self.text_v_align = 0
        self.text_justlast = 0
        self.rotation = 0

        self.fg_color = 0x00
        self.bg_color = 0x00
        self.pattern = 0
        self.has_fill = 0
        self.has_dxf_fill = 0
        self.fill_index = 0
        self.fill_count = 0

        self.border_index = 0
        self.has_border = 0
        self.has_dxf_border = 0
        self.border_count = 0

        self.bottom = 0
        self.bottom_color = 0x0
        self.diag_border = 0
        self.diag_color = 0x0
        self.diag_type = 0
        self.left = 0
        self.left_color = 0x0
        self.right = 0
        self.right_color = 0x0
        self.top = 0
        self.top_color = 0x0

        self.indent = 0
        self.shrink = 0
        self.merge_range = 0
        self.reading_order = 0
        self.just_distrib = 0
        self.color_indexed = 0
        self.font_only = 0

    ###########################################################################
    #
    # Private API.
    #
    ###########################################################################

    def _assemble_xml_file(self):
        # Assemble and write the XML file.

        # Write the XML declaration.
        self._xml_declaration()

        # Close the file.
        self._xml_close()

    def get_align_properties(self):
        # TODO. Temp method.
        return 0, 0

    def get_protection_properties(self):
        # TODO. Temp method.
        return []

    ###########################################################################
    #
    # XML methods.
    #
    ###########################################################################
