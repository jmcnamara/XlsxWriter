###############################################################################
#
# Comments - A class for writing the Excel XLSX Worksheet file.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2025, John McNamara, jmcnamara@cpan.org
#

from . import xmlwriter
from .utility import (
    _preserve_whitespace,
    _xl_color,
    xl_cell_to_rowcol,
    xl_rowcol_to_cell,
)


###########################################################################
#
# A comment type class.
#
###########################################################################
class CommentType:
    """
    A class to represent a comment in an Excel worksheet.

    """

    def __init__(self, row: int, col: int, text: str, options: dict = None):
        """
        Initialize a Comment instance.

        Args:
            row (int): The row number of the comment.
            col (int): The column number of the comment.
            text (str): The text of the comment.
            options (dict): Additional options for the comment.
        """
        self.row = row
        self.col = col
        self.text = text

        self.author = None
        self.color = "#ffffe1"

        self.start_row = 0
        self.start_col = 0

        self.is_visible = None

        self.width = 128
        self.height = 74

        self.x_scale = 1
        self.y_scale = 1
        self.x_offset = 0
        self.y_offset = 0

        self.font_size = 8
        self.font_name = "Tahoma"
        self.font_family = 2

        self.vertices = []

        # Set the default start cell and offsets for the comment.
        self.set_offsets(self.row, self.col)

        # Set any user supplied options.
        self._set_user_options(options)

    def _set_user_options(self, options=None):
        """
        This method handles the additional optional parameters to
        ``write_comment()``.
        """
        if options is None:
            return

        # Overwrite the defaults with any user supplied values. Incorrect or
        # misspelled parameters are silently ignored.
        self.width = options.get("width", self.width)
        self.author = options.get("author", self.author)
        self.height = options.get("height", self.height)
        self.x_offset = options.get("x_offset", self.x_offset)
        self.y_offset = options.get("y_offset", self.y_offset)
        self.start_col = options.get("start_col", self.start_col)
        self.start_row = options.get("start_row", self.start_row)
        self.font_size = options.get("font_size", self.font_size)
        self.font_name = options.get("font_name", self.font_name)
        self.is_visible = options.get("visible", self.is_visible)
        self.font_family = options.get("font_family", self.font_family)

        if options.get("color"):
            # Set the comment background color.
            self.color = _xl_color(options["color"]).lower()

            # Convert from Excel XML style color to XML html style color.
            self.color = options["color"].replace("ff", "#", 1)

        # Convert a cell reference to a row and column.
        if options.get("start_cell"):
            (start_row, start_col) = xl_cell_to_rowcol(options["start_cell"])
            self.start_row = start_row
            self.start_col = start_col

        # Scale the size of the comment box if required.
        if options.get("x_scale"):
            self.width = self.width * options["x_scale"]

        if options.get("y_scale"):
            self.height = self.height * options["y_scale"]

        # Round the dimensions to the nearest pixel.
        self.width = int(0.5 + self.width)
        self.height = int(0.5 + self.height)

    def set_offsets(self, row: int, col: int):
        """
        Set the default start cell and offsets for the comment. These are
        generally a fixed offset relative to the parent cell. However there are
        some edge cases for cells at the, well, edges.
        """
        row_max = 1048576
        col_max = 16384

        if self.row == 0:
            self.y_offset = 2
            self.start_row = 0
        elif self.row == row_max - 3:
            self.y_offset = 16
            self.start_row = row_max - 7
        elif self.row == row_max - 2:
            self.y_offset = 16
            self.start_row = row_max - 6
        elif self.row == row_max - 1:
            self.y_offset = 14
            self.start_row = row_max - 5
        else:
            self.y_offset = 10
            self.start_row = row - 1

        if self.col == col_max - 3:
            self.x_offset = 49
            self.start_col = col_max - 6
        elif self.col == col_max - 2:
            self.x_offset = 49
            self.start_col = col_max - 5
        elif self.col == col_max - 1:
            self.x_offset = 49
            self.start_col = col_max - 4
        else:
            self.x_offset = 15
            self.start_col = col + 1


###########################################################################
#
# The file writer class for the Excel XLSX Comments file.
#
###########################################################################
class Comments(xmlwriter.XMLwriter):
    """
    A class for writing the Excel XLSX Comments file.


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

        super().__init__()
        self.author_ids = {}

    ###########################################################################
    #
    # Private API.
    #
    ###########################################################################

    def _assemble_xml_file(self, comments_data=None):
        # Assemble and write the XML file.

        if comments_data is None:
            comments_data = []

        # Write the XML declaration.
        self._xml_declaration()

        # Write the comments element.
        self._write_comments()

        # Write the authors element.
        self._write_authors(comments_data)

        # Write the commentList element.
        self._write_comment_list(comments_data)

        self._xml_end_tag("comments")

        # Close the file.
        self._xml_close()

    ###########################################################################
    #
    # XML methods.
    #
    ###########################################################################

    def _write_comments(self):
        # Write the <comments> element.
        xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

        attributes = [("xmlns", xmlns)]

        self._xml_start_tag("comments", attributes)

    def _write_authors(self, comment_data):
        # Write the <authors> element.
        author_count = 0

        self._xml_start_tag("authors")

        for comment in comment_data:
            author = comment.author

            if author is not None and author not in self.author_ids:
                # Store the author id.
                self.author_ids[author] = author_count
                author_count += 1

                # Write the author element.
                self._write_author(author)

        self._xml_end_tag("authors")

    def _write_author(self, data):
        # Write the <author> element.
        self._xml_data_element("author", data)

    def _write_comment_list(self, comment_data):
        # Write the <commentList> element.
        self._xml_start_tag("commentList")

        for comment in comment_data:
            # Look up the author id.
            author_id = None
            if comment.author is not None:
                author_id = self.author_ids[comment.author]

            # Write the comment element.
            self._write_comment(comment, author_id)

        self._xml_end_tag("commentList")

    def _write_comment(self, comment: CommentType, author_id: int):
        # Write the <comment> element.
        ref = xl_rowcol_to_cell(comment.row, comment.col)

        attributes = [("ref", ref)]

        if author_id is not None:
            attributes.append(("authorId", author_id))

        self._xml_start_tag("comment", attributes)

        # Write the text element.
        self._write_text(comment)

        self._xml_end_tag("comment")

    def _write_text(self, comment: CommentType):
        # Write the <text> element.
        self._xml_start_tag("text")

        # Write the text r element.
        self._write_text_r(comment)

        self._xml_end_tag("text")

    def _write_text_r(self, comment: CommentType):
        # Write the <r> element.
        self._xml_start_tag("r")

        # Write the rPr element.
        self._write_r_pr(comment)

        # Write the text r element.
        self._write_text_t(comment.text)

        self._xml_end_tag("r")

    def _write_text_t(self, text):
        # Write the text <t> element.
        attributes = []

        if _preserve_whitespace(text):
            attributes.append(("xml:space", "preserve"))

        self._xml_data_element("t", text, attributes)

    def _write_r_pr(self, comment):
        # Write the <rPr> element.
        self._xml_start_tag("rPr")

        # Write the sz element.
        self._write_sz(comment.font_size)

        # Write the color element.
        self._write_color()

        # Write the rFont element.
        self._write_r_font(comment.font_name)

        # Write the family element.
        self._write_family(comment.font_family)

        self._xml_end_tag("rPr")

    def _write_sz(self, font_size: int):
        # Write the <sz> element.
        attributes = [("val", font_size)]

        self._xml_empty_tag("sz", attributes)

    def _write_color(self):
        # Write the <color> element.
        attributes = [("indexed", 81)]

        self._xml_empty_tag("color", attributes)

    def _write_r_font(self, font_name: str):
        # Write the <rFont> element.
        attributes = [("val", font_name)]

        self._xml_empty_tag("rFont", attributes)

    def _write_family(self, font_family: int):
        # Write the <family> element.
        attributes = [("val", font_family)]

        self._xml_empty_tag("family", attributes)
