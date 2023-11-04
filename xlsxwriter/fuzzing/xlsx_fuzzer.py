#!/usr/bin/python3
# Copyright 2023 Google LLC
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#      http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
#
################################################################################
from enum import Enum

import atheris
import sys
from io import BytesIO

from fuzz_helpers import EnhancedFuzzedDataProvider
import struct

with atheris.instrument_imports(include=['xlsxwriter']):
    import xlsxwriter
    import xlsxwriter.worksheet
    from xlsxwriter.exceptions import XlsxWriterException


class FuncChoice(Enum):
    WRITE_STRING = 0
    WRITE_NUMBER = 1
    WRITE_FORMULA = 2


choices = [FuncChoice.WRITE_STRING, FuncChoice.WRITE_NUMBER, FuncChoice.WRITE_FORMULA]


def TestOneInput(data):
    fdp = EnhancedFuzzedDataProvider(data)

    try:
        out = BytesIO()
        with xlsxwriter.Workbook(out) as wb:
            ws = wb.add_worksheet()

            data = fdp.ConsumeRandomString()
            func_choice = fdp.PickValueInList(choices)

            for row in range(fdp.ConsumeIntInRange(0, 500)):
                for col in range(fdp.ConsumeIntInRange(0, 500)):
                    if func_choice is FuncChoice.WRITE_STRING:
                        ws.write_string(row, col, data)
                    elif func_choice is FuncChoice.WRITE_NUMBER:
                        ws.write_number(row, col, data)
                    else:
                        ws.write_formula(row, col, data)
    except (XlsxWriterException, struct.error):
        return -1
    except TypeError as e:
        if 'must be real number' in str(e):
            return -1
        raise e


def main():
    atheris.Setup(sys.argv, TestOneInput)
    atheris.Fuzz()


if __name__ == "__main__":
    main()
