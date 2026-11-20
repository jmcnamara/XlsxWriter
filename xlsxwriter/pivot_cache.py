###############################################################################
#
# PivotCache - A class for writing the Excel XLSX pivotCacheDefinition file.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2026, John McNamara, jmcnamara@cpan.org
#

from typing import List, Tuple

from xlsxwriter import xmlwriter
from xlsxwriter.exceptions import PivotTableError
from xlsxwriter.utility import xl_range, xl_rowcol_to_cell


class PivotCache(xmlwriter.XMLwriter):
    """
    A class for writing the Excel XLSX pivotCacheDefinition file. The cache
    doesn't contain any records and is refreshed by Excel when the file is
    loaded.

    """

    ###########################################################################
    #
    # Public API.
    #
    ###########################################################################

    def __init__(self, source: Tuple[str, int, int, int, int]) -> None:
        """
        Constructor.

        """

        super().__init__()

        self.source = source
        self.field_names: List[str] = []

    ###########################################################################
    #
    # Private API.
    #
    ###########################################################################

    def _validate_field_names(self) -> None:
        # Validate that the header cells in the source data are non-blank and
        # unique, ignoring case, since Excel uses them as the pivot table field
        # names.
        sheet_name, first_row, first_col, _, _ = self.source
        seen_names: List[str] = []

        for col_offset, name in enumerate(self.field_names):
            cell = xl_rowcol_to_cell(first_row, first_col + col_offset)

            if not name:
                raise PivotTableError(
                    f"Header cell '{sheet_name}!{cell}' in the pivot table data "
                    f"source is blank."
                )

            if name.lower() in seen_names:
                raise PivotTableError(
                    f"Header '{name}' in cell '{sheet_name}!{cell}' is "
                    f"duplicated in the pivot table data source."
                )

            seen_names.append(name.lower())

    def _assemble_xml_file(self) -> None:
        # Assemble and write the XML file.

        # Write the XML declaration.
        self._xml_declaration()

        # Write the pivotCacheDefinition element.
        self._write_pivot_cache_definition()

        # Write the cacheSource element.
        self._write_cache_source()

        # Write the cacheFields element.
        self._write_cache_fields()

        # Close the pivotCacheDefinition tag.
        self._xml_end_tag("pivotCacheDefinition")

        # Close the file.
        self._xml_close()

    ###########################################################################
    #
    # XML methods.
    #
    ###########################################################################

    def _write_pivot_cache_definition(self) -> None:
        # Write the <pivotCacheDefinition> element.
        schema = "http://schemas.openxmlformats.org/"
        xmlns = schema + "spreadsheetml/2006/main"
        xmlns_r = schema + "officeDocument/2006/relationships"
        _, first_row, _, last_row, _ = self.source

        attributes = [
            ("xmlns", xmlns),
            ("xmlns:r", xmlns_r),
            ("saveData", 0),
            ("refreshOnLoad", 1),
            ("createdVersion", 3),
            ("refreshedVersion", 3),
            ("minRefreshableVersion", 3),
            ("recordCount", last_row - first_row),
        ]

        self._xml_start_tag("pivotCacheDefinition", attributes)

    def _write_cache_source(self) -> None:
        # Write the <cacheSource> element.
        attributes = [("type", "worksheet")]

        self._xml_start_tag("cacheSource", attributes)

        # Write the worksheetSource element.
        self._write_worksheet_source()

        self._xml_end_tag("cacheSource")

    def _write_worksheet_source(self) -> None:
        # Write the <worksheetSource> element.
        sheet_name, first_row, first_col, last_row, last_col = self.source

        attributes = [
            ("ref", xl_range(first_row, first_col, last_row, last_col)),
            ("sheet", sheet_name),
        ]

        self._xml_empty_tag("worksheetSource", attributes)

    def _write_cache_fields(self) -> None:
        # Write the <cacheFields> element.
        attributes = [("count", len(self.field_names))]

        self._xml_start_tag("cacheFields", attributes)

        for name in self.field_names:
            # Write the cacheField element.
            self._write_cache_field(name)

        self._xml_end_tag("cacheFields")

    def _write_cache_field(self, name: str) -> None:
        # Write the <cacheField> element.
        attributes = [
            ("name", name),
            ("numFmtId", 0),
        ]

        self._xml_start_tag("cacheField", attributes)

        # Write the sharedItems element.
        self._xml_empty_tag("sharedItems")

        self._xml_end_tag("cacheField")
