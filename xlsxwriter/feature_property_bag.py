###############################################################################
#
# FeaturePropertyBag - A class for writing the Excel XLSX FeaturePropertyBag file.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2013-2024, John McNamara, jmcnamara@cpan.org
#

from . import xmlwriter


class FeaturePropertyBag(xmlwriter.XMLwriter):
    """
    A class for writing the Excel XLSX FeaturePropertyBag file.


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
        self.bag_id_count = 0
        self.checkbox = False

    ###########################################################################
    #
    # Private API.
    #
    ###########################################################################

    def _assemble_xml_file(self):
        # Assemble and write the XML file.

        # Write the XML declaration.
        self._xml_declaration()

        # Write the featurePropertyBag element.
        self._write_feature_property_bag()

        # Close the file.
        self._xml_close()

    ###########################################################################
    #
    # XML methods.
    #
    ###########################################################################

    def _write_feature_property_bag(self):
        # Write the <FeaturePropertyBags> element.

        xmlns = (
            "http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag"
        )

        attributes = [
            ("xmlns", xmlns),
        ]

        self._xml_start_tag("FeaturePropertyBags", attributes)

        # Write the <bag> elements.
        if self.checkbox:
            self._write_checkbox()

        self._xml_end_tag("FeaturePropertyBags")

    def _write_checkbox(self):
        # Write a checkbox element.

        # add Checkbox element
        attributes1 = [
            ("type", "Checkbox"),
        ]
        self._xml_empty_tag("bag", attributes1)
        # add XFControls element
        attributes2 = [("type", "XFControls")]
        self._xml_start_tag("bag", attributes2)
        attributes3 = [("k", "CellControl")]
        self._xml_data_element("bagId", self.bag_id_count, attributes3)
        self.bag_id_count += 1
        self._xml_end_tag("bag")
        # add XFComplement element
        attributes4 = [("type", "XFComplement")]
        self._xml_start_tag("bag", attributes4)
        attributes5 = [("k", "XFControls")]
        self._xml_data_element("bagId", self.bag_id_count, attributes5)
        self.bag_id_count += 1
        self._xml_end_tag("bag")
        # add XFComplements element
        attributes6 = [
            ("type", "XFComplements"),
            ("extRef", "XFComplementsMapperExtRef"),
        ]
        self._xml_start_tag("bag", attributes6)
        attributes7 = [("k", "MappedFeaturePropertyBags")]
        self._xml_start_tag("a", attributes7)
        self._xml_data_element("bagId", self.bag_id_count)
        self.bag_id_count += 1
        self._xml_end_tag("a")
        self._xml_end_tag("bag")
