###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2024, John McNamara, jmcnamara@cpan.org
#

import unittest
from io import StringIO

from ...feature_property_bag import FeaturePropertyBag
from ..helperfunctions import _xml_to_list


class TestAssembleFeaturePropertyBag(unittest.TestCase):
    """
    Test assembling a complete FeaturePropertyBag file.

    """

    def test_assemble_xml_file(self):
        """Test for checkbox only."""
        self.maxDiff = None

        fh = StringIO()
        feature_property_bag = FeaturePropertyBag()
        feature_property_bag._set_filehandle(fh)

        feature_property_bag.checkbox = True

        feature_property_bag._assemble_xml_file()

        exp = _xml_to_list(
            """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <FeaturePropertyBags xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag">
                    <bag type="Checkbox"/>
                    <bag type="XFControls">
                        <bagId k="CellControl">0</bagId>
                    </bag>
                    <bag type="XFComplement">
                        <bagId k="XFControls">1</bagId>
                    </bag>
                    <bag type="XFComplements" extRef="XFComplementsMapperExtRef">
                        <a k="MappedFeaturePropertyBags">
                            <bagId>2</bagId>
                        </a>
                    </bag>
                </FeaturePropertyBags>
                """
        )

        got = _xml_to_list(fh.getvalue())

        self.assertEqual(got, exp)
