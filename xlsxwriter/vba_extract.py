#!python

##############################################################################
#
# vba_extract - A simple utility to extract a vbaProject.bin binary from an
# Excel 2007+ xlsm file for insertion into an XlsxWriter file.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2025, John McNamara, jmcnamara@cpan.org
#

import sys
from zipfile import BadZipFile, ZipFile


def extract_file(xlsm_zip, filename):
    """
    Extract a single file from an Excel xlsm macro file.

    :param xlsm_zip: The zip file to extract from
    :param filename: The file to extract
    """
    data = xlsm_zip.read("xl/" + filename)

    # Write the data to a local file.
    with open(filename, "wb") as file:
        file.write(data)


# The VBA project file and project signature file we want to extract.
VBA_FILENAME = "vbaProject.bin"
VBA_SIGNATURE_FILENAME = "vbaProjectSignature.bin"


def main():
    """
    vba_extract cli
    """
    # Get the xlsm file name from the commandline.
    if len(sys.argv) > 1:
        xlsm_file = sys.argv[1]
    else:
        print(
            "\nUtility to extract a vbaProject.bin binary from an Excel 2007+ "
            "xlsm macro file for insertion into an XlsxWriter file.\n"
            "If the macros are digitally signed, extracts also "
            "a vbaProjectSignature.bin file.\n"
            "\n"
            "See: https://xlsxwriter.readthedocs.io/working_with_macros.html\n"
            "\n"
            "Usage: vba_extract file.xlsm\n"
        )
        sys.exit()

    try:
        # Open the Excel xlsm file as a zip file.
        with ZipFile(xlsm_file, "r") as xlsm_zip:
            # Read the xl/vbaProject.bin file.
            extract_file(xlsm_zip, VBA_FILENAME)
            print(f"Extracted: {VBA_FILENAME}")

            if "xl/" + VBA_SIGNATURE_FILENAME in xlsm_zip.namelist():
                extract_file(xlsm_zip, VBA_SIGNATURE_FILENAME)
                print(f"Extracted: {VBA_SIGNATURE_FILENAME}")

    except IOError as e:
        print(f"File error: {str(e)}")
        sys.exit()

    except KeyError as e:
        # Usually when there isn't a xl/vbaProject.bin member in the file.
        print(f"File error: {str(e)}")
        print(f"File may not be an Excel xlsm macro file: '{xlsm_file}'")
        sys.exit()

    except BadZipFile as e:
        # Usually if the file is an xls file and not an xlsm file.
        print(f"File error: {str(e)}: '{xlsm_file}'")
        print("File may not be an Excel xlsm macro file.")
        sys.exit()

    except Exception as e:  # pylint: disable=broad-exception-caught
        # Catch any other exceptions.
        print(f"File error: {str(e)}")
        sys.exit()


if __name__ == "__main__":
    main()
