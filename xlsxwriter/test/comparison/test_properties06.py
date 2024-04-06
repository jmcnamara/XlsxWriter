###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparison_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):

        self.set_filename("properties06.xlsx")

    def test_create_file(self):
        """Test the creation of a simple XlsxWriter file."""

        workbook = Workbook(self.got_filename)

        worksheet = workbook.add_worksheet()

        company_guid = "2096f6a2-d2f7-48be-b329-b73aaa526e5d"
        site_id = "cb46c030-1825-4e81-a295-151c039dbf02"
        action_id = "88124cf5-1340-457d-90e1-0000a9427c99"

        workbook.set_custom_property(
            f"MSIP_Label_{company_guid}_Enabled", "true", "text"
        )
        workbook.set_custom_property(
            f"MSIP_Label_{company_guid}_SetDate", "2024-01-01T12:00:00Z", "text"
        )
        workbook.set_custom_property(
            f"MSIP_Label_{company_guid}_Method", "Privileged", "text"
        )
        workbook.set_custom_property(
            f"MSIP_Label_{company_guid}_Name", "Confidential", "text"
        )
        workbook.set_custom_property(
            f"MSIP_Label_{company_guid}_SiteId", site_id, "text"
        )
        workbook.set_custom_property(
            f"MSIP_Label_{company_guid}_ActionId", action_id, "text"
        )
        workbook.set_custom_property(
            f"MSIP_Label_{company_guid}_ContentBits", "2", "text"
        )

        workbook.close()

        self.assertExcelEqual()
