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

        self.bold = 0
        self.underline = 0
        self.italic = 0
        self.font_name = 'Calibri'
        self.font_size = 11
        self.font_color = 0x0
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
    # Format properties.
    #
    ###########################################################################

    def set_num_format(self, num_format):
        """
        Set the num_format property.

        Args:
            num_format: Default is 0.
            None. Defaults to turning property on.


        Returns:
            Nothing.

        """
        self.num_format = num_format

    def set_bold(self, bold):
        """
        Set the bold property.

        Args:
            bold: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.bold = bold

    def set_underline(self, underline):
        """
        Set the underline property.

        Args:
            underline: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.underline = underline

    def set_italic(self, italic):
        """
        Set the italic property.

        Args:
            italic: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.italic = italic

    def set_font_name(self, font_name):
        """
        Set the font_name property.

        Args:
            font_name: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.font_name = font_name

    def set_font_size(self, font_size):
        """
        Set the font_size property.

        Args:
            font_size: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.font_size = font_size

    def set_font_color(self, font_color):
        """
        Set the font_color property.

        Args:
            font_color: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.font_color = font_color

    def set_font_strikeout(self, font_strikeout):
        """
        Set the font_strikeout property.

        Args:
            font_strikeout: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.font_strikeout = font_strikeout

    def set_font_outline(self, font_outline):
        """
        Set the font_outline property.

        Args:
            font_outline: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.font_outline = font_outline

    def set_font_shadow(self, font_shadow):
        """
        Set the font_shadow property.

        Args:
            font_shadow: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.font_shadow = font_shadow

    def set_font_script(self, font_script):
        """
        Set the font_script property.

        Args:
            font_script: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.font_script = font_script

    def set_font_family(self, font_family):
        """
        Set the font_family property.

        Args:
            font_family: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.font_family = font_family

    def set_font_charset(self, font_charset):
        """
        Set the font_charset property.

        Args:
            font_charset: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.font_charset = font_charset

    def set_font_scheme(self, font_scheme):
        """
        Set the font_scheme property.

        Args:
            font_scheme: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.font_scheme = font_scheme

    def set_font_condense(self, font_condense):
        """
        Set the font_condense property.

        Args:
            font_condense: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.font_condense = font_condense

    def set_font_extend(self, font_extend):
        """
        Set the font_extend property.

        Args:
            font_extend: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.font_extend = font_extend

    def set_theme(self, theme):
        """
        Set the theme property.

        Args:
            theme: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.theme = theme

    def set_hyperlink(self, hyperlink):
        """
        Set the hyperlink property.

        Args:
            hyperlink: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.hyperlink = hyperlink

    def set_hidden(self, hidden):
        """
        Set the hidden property.

        Args:
            hidden: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.hidden = hidden

    def set_locked(self, locked):
        """
        Set the locked property.

        Args:
            locked: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.locked = locked

    def set_text_h_align(self, text_h_align):
        """
        Set the text_h_align property.

        Args:
            text_h_align: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.text_h_align = text_h_align

    def set_text_wrap(self, text_wrap):
        """
        Set the text_wrap property.

        Args:
            text_wrap: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.text_wrap = text_wrap

    def set_text_v_align(self, text_v_align):
        """
        Set the text_v_align property.

        Args:
            text_v_align: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.text_v_align = text_v_align

    def set_text_justlast(self, text_justlast):
        """
        Set the text_justlast property.

        Args:
            text_justlast: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.text_justlast = text_justlast

    def set_rotation(self, rotation):
        """
        Set the rotation property.

        Args:
            rotation: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.rotation = rotation

    def set_fg_color(self, fg_color):
        """
        Set the fg_color property.

        Args:
            fg_color: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.fg_color = fg_color

    def set_bg_color(self, bg_color):
        """
        Set the bg_color property.

        Args:
            bg_color: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.bg_color = bg_color

    def set_pattern(self, pattern):
        """
        Set the pattern property.

        Args:
            pattern: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.pattern = pattern

    def set_bottom(self, bottom):
        """
        Set the bottom property.

        Args:
            bottom: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.bottom = bottom

    def set_bottom_color(self, bottom_color):
        """
        Set the bottom_color property.

        Args:
            bottom_color: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.bottom_color = bottom_color

    def set_diag_border(self, diag_border):
        """
        Set the diag_border property.

        Args:
            diag_border: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.diag_border = diag_border

    def set_diag_color(self, diag_color):
        """
        Set the diag_color property.

        Args:
            diag_color: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.diag_color = diag_color

    def set_diag_type(self, diag_type):
        """
        Set the diag_type property.

        Args:
            diag_type: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.diag_type = diag_type

    def set_left(self, left):
        """
        Set the left property.

        Args:
            left: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.left = left

    def set_left_color(self, left_color):
        """
        Set the left_color property.

        Args:
            left_color: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.left_color = left_color

    def set_right(self, right):
        """
        Set the right property.

        Args:
            right: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.right = right

    def set_right_color(self, right_color):
        """
        Set the right_color property.

        Args:
            right_color: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.right_color = right_color

    def set_top(self, top):
        """
        Set the top property.

        Args:
            top: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.top = top

    def set_top_color(self, top_color):
        """
        Set the top_color property.

        Args:
            top_color: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.top_color = top_color

    def set_indent(self, indent):
        """
        Set the indent property.

        Args:
            indent: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.indent = indent

    def set_shrink(self, shrink):
        """
        Set the shrink property.

        Args:
            shrink: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.shrink = shrink

    # Backward compatibility. TODO.
    def set_name(self, font_name):
        #  TODO
        self.font_name = font_name

    def set_size(self, font_size):
        #  TODO
        self.font_size = font_size

    def set_color(self, font_color):
        #  TODO
        self.font_color = font_color

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
