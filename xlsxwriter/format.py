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

    def __init__(self, properties={}):
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

        # Convert properties in the constructor to method calls.
        for key, value in properties.items():
            getattr(self, 'set_' + key)(value)

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

    def set_bold(self, bold=1):
        """
        Set the bold property.

        Args:
            bold: Default is 1.

        Returns:
            Nothing.

        """
        self.bold = bold

    def set_underline(self, underline=1):
        """
        Set the underline property.

        Args:
            underline: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.underline = underline

    def set_italic(self, italic=1):
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

    def set_font_strikeout(self, font_strikeout=1):
        """
        Set the font_strikeout property.

        Args:
            font_strikeout: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.font_strikeout = font_strikeout

    def set_font_outline(self, font_outline=1):
        """
        Set the font_outline property.

        Args:
            font_outline: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.font_outline = font_outline

    def set_font_shadow(self, font_shadow=1):
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

    def set_text_wrap(self, text_wrap=1):
        """
        Set the text_wrap property.

        Args:
            text_wrap: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.text_wrap = text_wrap

    def set_hidden(self, hidden=1):
        """
        Set the hidden property.

        Args:
            hidden: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.hidden = hidden

    def set_locked(self, locked=1):
        """
        Set the locked property.

        Args:
            locked: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.locked = locked

    def set_align(self, location):
        """
        Set the cell alignment.

        """
        location = location.lower()

        # Set horizontal alignment properties.
        if location == 'left':
            self.set_text_h_align(1)
        if location == 'centre':
            self.set_text_h_align(2)
        if location == 'center':
            self.set_text_h_align(2)
        if location == 'right':
            self.set_text_h_align(3)
        if location == 'fill':
            self.set_text_h_align(4)
        if location == 'justify':
            self.set_text_h_align(5)
        if location == 'center_across':
            self.set_text_h_align(6)
        if location == 'centre_across':
            self.set_text_h_align(6)
        if location == 'distributed':
            self.set_text_h_align(7)
        if location == 'justify_distributed':
            self.set_text_h_align(7)

        if location == 'justify_distributed':
            self.just_distrib = 1

        # Set vertical alignment properties.
        if location == 'top':
            self.set_text_v_align(1)
        if location == 'vcentre':
            self.set_text_v_align(2)
        if location == 'vcenter':
            self.set_text_v_align(2)
        if location == 'bottom':
            self.set_text_v_align(3)
        if location == 'vjustify':
            self.set_text_v_align(4)
        if location == 'vdistributed':
            self.set_text_v_align(5)

    def set_text_justlast(self, text_justlast=1):
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
        rotation = int(rotation)

        # Map user angle to Excel angle.
        if rotation == 270:
            rotation = 255
        elif rotation >= -90 or rotation <= 90:
            if rotation < 0:
                rotation = -rotation + 90
        else:
            raise Exception(
                "Rotation rotation outside range: -90 <= angle <= 90")

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

    def set_indent(self, indent=1):
        """
        Set the indent property.

        Args:
            indent: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.indent = indent

    def set_shrink(self, shrink=1):
        """
        Set the shrink property.

        Args:
            shrink: Default is 0.
            None. Defaults to turning property on.

        Returns:
            Nothing.

        """
        self.shrink = shrink

    ###########################################################################
    #
    # Internal Format properties. These aren't documented since they are
    # either only used internally or else are unlikely to be set by the user.
    #
    ###########################################################################

    def set_TODO_XXXXXXX(self):
        pass

    def set_has_font(self, has_font=1):
        # Set the has_font property.
        self.has_font = has_font

    def set_font_index(self, font_index):
        # Set the font_index property.
        self.font_index = font_index

    def set_num_format_index(self, num_format_index):
        # Set the num_format_index property.
        self.num_format_index = num_format_index

    def set_text_h_align(self, text_h_align):
        # Set the text_h_align property.
        self.text_h_align = text_h_align

    def set_text_v_align(self, text_v_align):
        # Set the text_v_align property.
        self.text_v_align = text_v_align

    def set_reading_order(self, reading_order=1):
        # Set the reading_order property.
        self.reading_order = reading_order

    # Compatibility methods.
    def set_name(self, font_name):
        #  For compatibility with Excel::Writer::XLSX.
        self.font_name = font_name

    def set_size(self, font_size):
        #  For compatibility with Excel::Writer::XLSX.
        self.font_size = font_size

    def set_color(self, font_color):
        #  For compatibility with Excel::Writer::XLSX.
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

    # Return properties for an Style xf <alignment> sub-element.
    def get_align_properties(self):
        """
        TODO

        """
        # Attributes to return
        changed = 0
        align = []

        # Check if any alignment options in the format have been changed.
        if (self.text_h_align or self.text_v_align or self.indent
                or self.rotation or self.text_wrap or self.shrink
                or self.reading_order):
            changed = 1
        else:
            return changed, align

        # Indent is only allowed for horizontal left, right and distributed.
        # If it is defined for any other alignment or no alignment has
        # been set then default to left alignment.
        if (self.indent
                and self.text_h_align != 1
                and self.text_h_align != 3
                and self.text_h_align != 7):
            self.text_h_align = 1

        # Check for properties that are mutually exclusive.
        if self.text_wrap:
            self.shrink = 0
        if self.text_h_align == 4:
            self.shrink = 0
        if self.text_h_align == 5:
            self.shrink = 0
        if self.text_h_align == 7:
            self.shrink = 0
        if self.text_h_align != 7:
            self.just_distrib = 0
        if self.indent:
            self.just_distrib = 0

        continuous = 'centerContinuous'

        if self.text_h_align == 1:
            align.append(('horizontal', 'left'))
        if self.text_h_align == 2:
            align.append(('horizontal', 'center'))
        if self.text_h_align == 3:
            align.append(('horizontal', 'right'))
        if self.text_h_align == 4:
            align.append(('horizontal', 'fill'))
        if self.text_h_align == 5:
            align.append(('horizontal', 'justify'))
        if self.text_h_align == 6:
            align.append(('horizontal', continuous))
        if self.text_h_align == 7:
            align.append(('horizontal', 'distributed'))

        if self.just_distrib:
            align.append(('justifyLastLine', 1))

        # Property 'vertical' => 'bottom' is a default. It sets applyAlignment
        # without an alignment sub-element.
        if self.text_v_align == 1:
            align.append(('vertical', 'top'))
        if self.text_v_align == 2:
            align.append(('vertical', 'center'))
        if self.text_v_align == 4:
            align.append(('vertical', 'justify'))
        if self.text_v_align == 5:
            align.append(('vertical', 'distributed'))

        if self.indent:
            align.append(('indent', self.indent))
        if self.rotation:
            align.append(('textRotation', self.rotation))

        if self.text_wrap:
            align.append(('wrapText', 1))
        if self.shrink:
            align.append(('shrinkToFit', 1))

        if self.reading_order == 1:
            align.append(('readingOrder', 1))
        if self.reading_order == 2:
            align.append(('readingOrder', 2))

        return changed, align

    def get_protection_properties(self):
        # TODO.
        attribs = []

        if not self.locked:
            attribs.append(('locked', 0))
        if self.hidden:
            attribs.append(('hidden', 1))

        return attribs

    def get_font_key(self):
        # Returns a unique hash key for a font. Used by Workbook.
        key = ':'.join(str(x) for x in (
            self.bold,
            self.color,
            self.font_charset,
            self.font_family,
            self.font_outline,
            self.font_script,
            self.font_shadow,
            self.font_strikeout,
            self.font_name,
            self.italic,
            self.font_size,
            self.underline))

        return key

    def get_border_key(self):
        # Returns a unique hash key for a border style. Used by Workbook.
        key = ':'.join(str(x) for x in (
            self.bottom,
            self.bottom_color,
            self.diag_border,
            self.diag_color,
            self.diag_type,
            self.left,
            self.left_color,
            self.right,
            self.right_color,
            self.top,
            self.top_color))

        return key

    def get_fill_key(self):
        # Returns a unique hash key for a fill style. Used by Workbook.
        key = ':'.join(str(x) for x in (
            self.pattern,
            self.bg_color,
            self.fg_color))

        return key

    def get_alignment_key(self):
        # Returns a unique hash key for alignment formats.

        key = ':'.join(str(x) for x in (
            self.text_h_align,
            self.text_v_align,
            self.indent,
            self.rotation,
            self.text_wrap,
            self.shrink,
            self.reading_order))

        return key

    ###########################################################################
    #
    # XML methods.
    #
    ###########################################################################
