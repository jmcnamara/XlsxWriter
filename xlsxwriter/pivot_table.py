###############################################################################
#
# PivotTable - A class for writing the Excel XLSX pivotTable file.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2026, John McNamara, jmcnamara@cpan.org
#

import re
from typing import Any, Dict, List, Optional, Sequence, Tuple, Union

from xlsxwriter import xmlwriter
from xlsxwriter.exceptions import PivotTableError
from xlsxwriter.utility import xl_cell_to_rowcol, xl_rowcol_to_cell

# A pivot table data source as a (sheet_name, first_row, first_col, last_row,
# last_col) tuple.
PivotTableSource = Tuple[str, int, int, int, int]

# The Excel worksheet row and column limits.
ROW_MAX = 1048576
COL_MAX = 16384

# Match a cell range like "A1:D9" after any "$" characters are removed.
RANGE_RE = re.compile(r"^([A-Z]{1,3}[1-9]\d*):([A-Z]{1,3}[1-9]\d*)$")

# Match a pivot table style name like "PivotStyleLight16" after any spaces
# are removed.
STYLE_RE = re.compile(r"^PivotStyle(Light|Medium|Dark)([1-9]|1\d|2[0-8])$")

# The valid add_pivot_table() options.
VALID_OPTIONS = {
    "source",
    "rows",
    "columns",
    "filters",
    "values",
    "style",
    "layout",
    "row_grand_totals",
    "column_grand_totals",
    "field_names",
}

# The valid value field options.
VALID_VALUE_OPTIONS = {"field", "function", "name", "num_format"}

# The valid pivot table layouts.
VALID_LAYOUTS = {"compact", "outline", "tabular"}

# The value field functions mapped to their dataField "subtotal" attribute
# value and the function name used in the default value field caption, such as
# "Sum of Volume".
VALUE_FUNCTIONS = {
    "sum": ("sum", "Sum"),
    "count": ("count", "Count"),
    "average": ("average", "Average"),
    "max": ("max", "Max"),
    "min": ("min", "Min"),
    "product": ("product", "Product"),
    "count_numbers": ("countNums", "Count"),
    "std_dev": ("stdDev", "StdDev"),
    "std_devp": ("stdDevp", "StdDevp"),
    "var": ("var", "Var"),
    "varp": ("varp", "Varp"),
}


class PivotTable(xmlwriter.XMLwriter):
    """
    A class for writing the Excel XLSX pivotTable file. The pivot table is
    written without any calculated data and is refreshed by Excel when the
    file is loaded.

    """

    ###########################################################################
    #
    # Public API.
    #
    ###########################################################################

    def __init__(self, row: int, col: int, options: Dict[str, Any]) -> None:
        """
        Constructor.

        """

        super().__init__()

        self.row = row
        self.col = col
        self.name = ""
        self.cache_id = 0
        self.field_names: List[str] = []
        self.source_field_names: List[str] = []
        self.row_fields: List[str] = []
        self.column_fields: List[str] = []
        self.filter_fields: List[str] = []
        self.value_fields: List[Dict[str, Any]] = []
        self.style: Optional[str] = "PivotStyleLight16"
        self.layout = "compact"
        self.show_row_grand_totals = True
        self.show_column_grand_totals = True

        for option in options:
            if option not in VALID_OPTIONS:
                raise PivotTableError(f"Unknown pivot table option '{option}'.")

        if "source" not in options:
            raise PivotTableError(
                "The pivot table data source must be set with the 'source' option."
            )

        self.source = _parse_source(options["source"])
        self.row_fields = _get_field_list(options, "rows")
        self.column_fields = _get_field_list(options, "columns")
        self.filter_fields = _get_field_list(options, "filters")
        self.value_fields = _get_value_fields(options)

        if "style" in options:
            self.style = _get_style(options["style"])

        if "layout" in options:
            self.layout = options["layout"]
            if self.layout not in VALID_LAYOUTS:
                raise PivotTableError(
                    f"Unknown pivot table layout '{self.layout}'. Valid layouts "
                    f"are 'compact', 'outline' and 'tabular'."
                )

        self.show_row_grand_totals = bool(options.get("row_grand_totals", True))
        self.show_column_grand_totals = bool(options.get("column_grand_totals", True))

        # Field names that the user can set for constant_memory sources, where
        # the header row has already been written to disk and can't be read.
        self.source_field_names = _get_field_list(options, "field_names")

        self._validate()

    ###########################################################################
    #
    # Private API.
    #
    ###########################################################################

    def _assemble_xml_file(self) -> None:
        # Assemble and write the XML file.

        # Write the XML declaration.
        self._xml_declaration()

        # Write the pivotTableDefinition element.
        self._write_pivot_table_definition()

        # Write the location element.
        self._write_location()

        # Write the pivotFields element.
        self._write_pivot_fields()

        has_fields = bool(self.row_fields or self.column_fields or self.value_fields)

        if self.row_fields:
            # Write the rowFields element.
            self._write_row_fields()

        if has_fields and (not self.row_fields or self.show_column_grand_totals):
            # Write the rowItems element.
            self._write_axis_items("rowItems", bool(self.row_fields))

        # Multiple value fields are displayed in columns via a virtual
        # "Values" field that is added to the column fields.
        has_values_field = len(self.value_fields) > 1

        if self.column_fields or has_values_field:
            # Write the colFields element.
            self._write_col_fields(has_values_field)

        if has_fields and (not self.column_fields or self.show_row_grand_totals):
            if has_values_field:
                # Write the colItems element for multiple value fields.
                self._write_value_col_items()
            else:
                # Write the colItems element.
                self._write_axis_items("colItems", bool(self.column_fields))

        if self.filter_fields:
            # Write the pageFields element.
            self._write_page_fields()

        if self.value_fields:
            # Write the dataFields element.
            self._write_data_fields()

        # Write the pivotTableStyleInfo element.
        self._write_pivot_table_style_info()

        # Close the pivotTableDefinition tag.
        self._xml_end_tag("pivotTableDefinition")

        # Close the file.
        self._xml_close()

    def _validate(self) -> None:
        # Validate the pivot table options:
        #
        # - The source must have at least one data row.
        # - Any field names must match the number of source columns.
        # - A field can only be used once in the row, column and filter fields,
        #   but it can also be a value field.
        # - Custom value field names must be unique, ignoring case.
        #
        _, first_row, _, last_row, _ = self.source

        if first_row == last_row:
            raise PivotTableError(
                "The pivot table data source must have a header row and at "
                "least one data row."
            )

        if self.source_field_names and len(self.source_field_names) != (
            self._num_fields()
        ):
            raise PivotTableError(
                f"The number of field names ({len(self.source_field_names)}) "
                f"doesn't match the number of data source columns "
                f"({self._num_fields()})."
            )

        seen_names: List[str] = []
        for name in self.row_fields + self.column_fields + self.filter_fields:
            if name in seen_names:
                raise PivotTableError(
                    f"Field '{name}' is used more than once in the row, column "
                    f"or filter fields."
                )
            seen_names.append(name)

        value_names = [name.lower() for name in self._value_field_names()]
        for value_field, name in zip(self.value_fields, self._value_field_names()):
            if value_field["name"] and value_names.count(name.lower()) > 1:
                raise PivotTableError(
                    f"Value field name '{name}' is used more than once."
                )

    def _validate_field_names(self) -> None:
        # Validate that the pivot table field names match the field names in
        # the header row of the source data. This can only be done when the
        # workbook is saved since the header data may not have been written
        # before then. A custom value field name also can't be the same as a
        # header name, ignoring case.
        value_names = [value_field["field"] for value_field in self.value_fields]

        for name in (
            self.row_fields + self.column_fields + self.filter_fields + value_names
        ):
            if self._field_index(name) is None:
                raise PivotTableError(
                    f"Field '{name}' doesn't match a header in the source data."
                )

        header_names = [name.lower() for name in self.field_names]
        for value_field in self.value_fields:
            name = value_field["name"]
            if name and name.lower() in header_names:
                raise PivotTableError(
                    f"Value field name '{name}' is the same as a header in the "
                    f"source data."
                )

    def _set_num_format_indices(self, xf_formats: List[Any]) -> None:
        # Set the number format index of any value field number formats from
        # the prepared workbook cell formats.
        for value_field in self.value_fields:
            if "xf_index" in value_field:
                xf_format = xf_formats[value_field["xf_index"]]
                value_field["num_format_index"] = xf_format.num_format_index

    def _num_fields(self) -> int:
        # The number of fields (columns) in the source data.
        _, _, first_col, _, last_col = self.source
        return last_col - first_col + 1

    def _body_row(self) -> int:
        # Get the first row of the pivot table body. The (row, col) position
        # is the top left of the pivot table including any filter fields. The
        # body is below the filter fields and a blank row.
        if self.filter_fields:
            return self.row + len(self.filter_fields) + 1

        return self.row

    def _field_index(self, name: str) -> Optional[int]:
        # Get the index of a field in the source data from its header name.
        if name in self.field_names:
            return self.field_names.index(name)

        return None

    def _value_field_names(self) -> List[str]:
        # Get the display names of the value fields. These are either the
        # custom name or a default name like "Sum of Volume". Excel makes
        # duplicate default names unique by appending a number, such as "Sum
        # of Volume2". The names are compared ignoring case.
        names: List[str] = []
        seen: List[str] = []

        for value_field in self.value_fields:
            if value_field["name"]:
                name = value_field["name"]
            else:
                caption = VALUE_FUNCTIONS[value_field["function"]][1]
                name = f"{caption} of {value_field['field']}"

            if not value_field["name"] and name.lower() in seen:
                suffix = 2
                while f"{name}{suffix}".lower() in seen:
                    suffix += 1
                name = f"{name}{suffix}"

            names.append(name)
            seen.append(name.lower())

        return names

    ###########################################################################
    #
    # XML methods.
    #
    ###########################################################################

    def _write_pivot_table_definition(self) -> None:
        # Write the <pivotTableDefinition> element.
        schema = "http://schemas.openxmlformats.org/"
        xmlns = schema + "spreadsheetml/2006/main"

        attributes = [
            ("xmlns", xmlns),
            ("name", self.name),
            ("cacheId", self.cache_id),
            ("applyNumberFormats", 0),
            ("applyBorderFormats", 0),
            ("applyFontFormats", 0),
            ("applyPatternFormats", 0),
            ("applyAlignmentFormats", 0),
            ("applyWidthHeightFormats", 1),
            ("dataCaption", "Values"),
            ("updatedVersion", 3),
            ("minRefreshableVersion", 3),
            ("showCalcMbrs", 0),
            ("useAutoFormatting", 1),
        ]

        # The Excel UI and XML names are reversed. The UI row grand totals are
        # the "Grand Total" column, which is the XML colGrandTotals, and vice
        # versa.
        if not self.show_column_grand_totals:
            attributes.append(("rowGrandTotals", 0))

        if not self.show_row_grand_totals:
            attributes.append(("colGrandTotals", 0))

        attributes.extend(
            [
                ("itemPrintTitles", 1),
                ("createdVersion", 3),
                ("indent", 0),
            ]
        )

        if self.layout == "compact":
            attributes.append(("outline", 1))
            attributes.append(("outlineData", 1))
        elif self.layout == "outline":
            attributes.append(("compact", 0))
            attributes.append(("outline", 1))
            attributes.append(("outlineData", 1))
            attributes.append(("compactData", 0))
        else:
            attributes.append(("compact", 0))
            attributes.append(("compactData", 0))

        attributes.append(("multipleFieldFilters", 0))

        self._xml_start_tag("pivotTableDefinition", attributes)

    def _write_location(self) -> None:
        # Write the <location> element.

        # There is a header row for each column field, including the virtual
        # "Values" field used for multiple value fields.
        num_col_fields = len(self.column_fields)
        if len(self.value_fields) > 1:
            num_col_fields += 1

        first_header_row = 1
        first_data_row = 1 + num_col_fields

        # In the compact layout the row fields share a column. In the outline
        # and tabular layouts each row field has its own column.
        if not self.row_fields and not self.column_fields:
            first_data_col = 0
        elif self.layout == "compact":
            first_data_col = 1
        else:
            first_data_col = max(len(self.row_fields), 1)

        # A pivot table with only filter fields has no header or data rows.
        if (
            not self.row_fields
            and not self.column_fields
            and not self.value_fields
            and self.filter_fields
        ):
            first_header_row = 0
            first_data_row = 0

        attributes = [
            ("ref", xl_rowcol_to_cell(self._body_row(), self.col)),
            ("firstHeaderRow", first_header_row),
            ("firstDataRow", first_data_row),
            ("firstDataCol", first_data_col),
        ]

        if self.filter_fields:
            attributes.append(("rowPageCount", len(self.filter_fields)))
            attributes.append(("colPageCount", 1))

        self._xml_empty_tag("location", attributes)

    def _write_pivot_fields(self) -> None:
        # Write the <pivotFields> element.
        count = self._num_fields()
        attributes = [("count", count)]

        self._xml_start_tag("pivotFields", attributes)

        for index in range(count):
            # Write the pivotField element.
            self._write_pivot_field(index)

        self._xml_end_tag("pivotFields")

    def _write_pivot_field(self, index: int) -> None:
        # Write the <pivotField> element.
        def is_field(names: List[str]) -> bool:
            return any(self._field_index(name) == index for name in names)

        is_row_field = is_field(self.row_fields)
        is_column_field = is_field(self.column_fields)
        is_filter_field = is_field(self.filter_fields)
        is_data_field = is_field([field["field"] for field in self.value_fields])

        attributes = []

        if is_row_field:
            attributes.append(("axis", "axisRow"))
        elif is_column_field:
            attributes.append(("axis", "axisCol"))
        elif is_filter_field:
            attributes.append(("axis", "axisPage"))

        if is_data_field:
            attributes.append(("dataField", 1))

        if self.layout == "outline":
            attributes.append(("compact", 0))
        elif self.layout == "tabular":
            attributes.append(("compact", 0))
            attributes.append(("outline", 0))

        attributes.append(("showAll", 0))

        if is_row_field or is_column_field or is_filter_field:
            self._xml_start_tag("pivotField", attributes)

            # Write the items element.
            self._write_items()

            self._xml_end_tag("pivotField")
        else:
            self._xml_empty_tag("pivotField", attributes)

    def _write_items(self) -> None:
        # Write the <items> element. The items are generated by Excel when the
        # pivot table is refreshed so we only write the default item.
        attributes = [("count", 1)]

        self._xml_start_tag("items", attributes)

        # Write the item element.
        self._xml_empty_tag("item", [("t", "default")])

        self._xml_end_tag("items")

    def _write_row_fields(self) -> None:
        # Write the <rowFields> element.
        attributes = [("count", len(self.row_fields))]

        self._xml_start_tag("rowFields", attributes)

        for name in self.row_fields:
            # Write the field element.
            self._xml_empty_tag("field", [("x", self._field_index(name))])

        self._xml_end_tag("rowFields")

    def _write_col_fields(self, has_values_field: bool) -> None:
        # Write the <colFields> element, with the virtual "Values" field
        # (index -2) after the column fields if there are multiple value
        # fields.
        indices = [self._field_index(name) for name in self.column_fields]

        if has_values_field:
            indices.append(-2)

        attributes = [("count", len(indices))]

        self._xml_start_tag("colFields", attributes)

        for index in indices:
            # Write the field element.
            self._xml_empty_tag("field", [("x", index)])

        self._xml_end_tag("colFields")

    def _write_value_col_items(self) -> None:
        # Write the <colItems> element for multiple value fields. If there are
        # no column fields Excel writes an item for each value field.
        # Otherwise we write a grand total for each value field.
        count = len(self.value_fields)
        has_column_fields = bool(self.column_fields)
        attributes = [("count", count)]

        self._xml_start_tag("colItems", attributes)

        for index in range(count):
            attributes = []

            if has_column_fields:
                attributes.append(("t", "grand"))

            if index > 0:
                attributes.append(("i", index))

            # Write the i element.
            self._xml_start_tag("i", attributes)

            # Write the x element.
            if has_column_fields or index == 0:
                self._xml_empty_tag("x")
            else:
                self._xml_empty_tag("x", [("v", index)])

            self._xml_end_tag("i")

        self._xml_end_tag("colItems")

    def _write_axis_items(self, tag: str, has_fields: bool) -> None:
        # Write the <rowItems> or <colItems> element. The items are generated
        # by Excel when the pivot table is refreshed so we only write the
        # grand total, if the axis has fields, or an empty item.
        attributes = [("count", 1)]

        self._xml_start_tag(tag, attributes)

        # Write the i element.
        if has_fields:
            self._xml_start_tag("i", [("t", "grand")])
            self._xml_empty_tag("x")
            self._xml_end_tag("i")
        else:
            self._xml_empty_tag("i")

        self._xml_end_tag(tag)

    def _write_page_fields(self) -> None:
        # Write the <pageFields> element.
        attributes = [("count", len(self.filter_fields))]

        self._xml_start_tag("pageFields", attributes)

        for name in self.filter_fields:
            # Write the pageField element.
            attributes = [("fld", self._field_index(name)), ("hier", -1)]
            self._xml_empty_tag("pageField", attributes)

        self._xml_end_tag("pageFields")

    def _write_data_fields(self) -> None:
        # Write the <dataFields> element.
        attributes = [("count", len(self.value_fields))]

        self._xml_start_tag("dataFields", attributes)

        names = self._value_field_names()
        for value_field, name in zip(self.value_fields, names):
            # Write the dataField element.
            self._write_data_field(value_field, name)

        self._xml_end_tag("dataFields")

    def _write_data_field(self, value_field: Dict[str, Any], name: str) -> None:
        # Write the <dataField> element.
        attributes = [
            ("name", name),
            ("fld", self._field_index(value_field["field"])),
        ]

        if value_field["function"] != "sum":
            attributes.append(("subtotal", VALUE_FUNCTIONS[value_field["function"]][0]))

        attributes.append(("baseField", 0))
        attributes.append(("baseItem", 0))

        if value_field["num_format_index"]:
            attributes.append(("numFmtId", value_field["num_format_index"]))

        self._xml_empty_tag("dataField", attributes)

    def _write_pivot_table_style_info(self) -> None:
        # Write the <pivotTableStyleInfo> element.
        attributes = []

        if self.style:
            attributes.append(("name", self.style))

        attributes.extend(
            [
                ("showRowHeaders", 1),
                ("showColHeaders", 1),
                ("showRowStripes", 0),
                ("showColStripes", 0),
                ("showLastColumn", 1),
            ]
        )

        self._xml_empty_tag("pivotTableStyleInfo", attributes)


###############################################################################
#
# Helper functions.
#
###############################################################################


def _parse_source(source: Union[str, Sequence[Any]]) -> PivotTableSource:
    # Convert a pivot table data source, either a range string like
    # "Sheet1!A1:D9" or a [sheet_name, first_row, first_col, last_row,
    # last_col] list, to a source tuple.
    if isinstance(source, str):
        parsed = _parse_source_range(source)
        if parsed is None:
            raise PivotTableError(
                f"The pivot table data source '{source}' isn't a valid range "
                f"like 'Sheet1!A1:D9'."
            )
    elif (
        isinstance(source, (list, tuple))
        and len(source) == 5
        and isinstance(source[0], str)
        and all(isinstance(value, int) for value in source[1:])
    ):
        parsed = (source[0], source[1], source[2], source[3], source[4])
    else:
        raise PivotTableError(
            f"The pivot table data source {_source_repr(source)} must be a "
            f"range string like 'Sheet1!A1:D9' or a list like "
            f"['Sheet1', 0, 0, 8, 3]."
        )

    _check_source_range(source, parsed)

    return parsed


def _check_source_range(source: Any, parsed: PivotTableSource) -> None:
    # Check that the source range is within Excel's limits and in order.
    _, first_row, first_col, last_row, last_col = parsed

    if (
        min(first_row, first_col, last_row, last_col) < 0
        or max(first_row, last_row) >= ROW_MAX
        or max(first_col, last_col) >= COL_MAX
    ):
        raise PivotTableError(
            f"The pivot table data source {_source_repr(source)} is outside "
            f"Excel's row or column limits."
        )

    if first_row > last_row or first_col > last_col:
        raise PivotTableError(
            f"The pivot table data source {_source_repr(source)} has its first "
            f"row or column after its last row or column."
        )


def _source_repr(source: Any) -> str:
    # Format a user data source for an error message.
    if isinstance(source, str):
        return f"'{source}'"

    return str(source)


def _parse_source_range(range_string: str) -> Optional[PivotTableSource]:
    # Parse a range string like "Sheet1!A1:D9" into a source tuple. Returns None
    # if the range is invalid.
    if range_string.startswith("="):
        range_string = range_string[1:]

    pos = range_string.rfind("!")
    if pos <= 0:
        return None

    sheet_name = range_string[:pos]
    cell_range = range_string[pos + 1 :].replace("$", "").upper()

    # Remove the quotes from a quoted sheet name and unescape any quotes.
    if len(sheet_name) > 1 and sheet_name[0] == "'" and sheet_name[-1] == "'":
        sheet_name = sheet_name[1:-1].replace("''", "'")

    if not sheet_name:
        return None

    match = RANGE_RE.match(cell_range)
    if not match:
        return None

    first_row, first_col = xl_cell_to_rowcol(match.group(1))
    last_row, last_col = xl_cell_to_rowcol(match.group(2))

    return (sheet_name, first_row, first_col, last_row, last_col)


def _get_field_list(options: Dict[str, Any], option: str) -> List[str]:
    # Get a list of field names, such as the "rows" fields, from the options.
    fields = options.get(option, [])

    if not isinstance(fields, (list, tuple)) or not all(
        isinstance(field, str) for field in fields
    ):
        raise PivotTableError(
            f"The pivot table '{option}' option must be a list of field names."
        )

    return list(fields)


def _get_value_fields(options: Dict[str, Any]) -> List[Dict[str, Any]]:
    # Get the list of value field dicts from the options.
    values = options.get("values", [])

    if not isinstance(values, (list, tuple)):
        raise PivotTableError(
            "The pivot table 'values' option must be a list of value field dicts."
        )

    value_fields = []

    for value in values:
        if not isinstance(value, dict) or not isinstance(value.get("field"), str):
            raise PivotTableError(
                f"Pivot table value field '{value}' must be a dict with a "
                f"'field' name like {{'field': 'Volume'}}."
            )

        for option in value:
            if option not in VALID_VALUE_OPTIONS:
                raise PivotTableError(
                    f"Unknown pivot table value field option '{option}'."
                )

        function = value.get("function", "sum")
        if function not in VALUE_FUNCTIONS:
            raise PivotTableError(
                f"Unknown pivot table value field function '{function}'."
            )

        value_fields.append(
            {
                "field": value["field"],
                "function": function,
                "name": value.get("name", ""),
                "num_format": value.get("num_format"),
                "num_format_index": 0,
            }
        )

    return value_fields


def _get_style(style: Optional[str]) -> Optional[str]:
    # Convert a style name like "Pivot Style Light 16" to the Excel style name
    # "PivotStyleLight16". A style of "None" or None turns off the style.
    if style is None or style == "None":
        return None

    excel_style = style.replace(" ", "")
    if not STYLE_RE.match(excel_style):
        raise PivotTableError(
            f"Unknown pivot table style '{style}'. Use a style like "
            f"'Pivot Style Light 16' or 'None'."
        )

    return excel_style
