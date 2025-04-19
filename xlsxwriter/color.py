###############################################################################
#
# Color - A class to represent Excel colors.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2025, John McNamara, jmcnamara@cpan.org
#


from enum import Enum
from typing import Tuple, Union


class ColorTypes(Enum):
    """
    Enum to represent different types of URLS.
    """

    RGB = 1
    THEME = 2


class Color:
    """
    A class to represent an Excel color.

    """

    def __init__(self, color: Union[str, int, Tuple[int, int]]):
        """
        Initialize a Color instance.

        Args:
            value (Union[str, int, Tuple[int, int]]): The value of the color
            (e.g., a hex string, an integer, or a tuple of two integers).
        """
        self._rgb_value: int = 0x000000
        self._type: ColorTypes = ColorTypes.RGB
        self._theme_color: Tuple[int, int] = (0, 0)
        self._is_automatic: bool = False

        if isinstance(color, str):
            self._parse_string_color(color)
            self._type = ColorTypes.RGB
        elif isinstance(color, int):
            if color > 0xFFFFFF:
                raise ValueError("RGB color must be in the range 0x000000 - 0xFFFFFF.")

            self._rgb_value = color
            self._type = ColorTypes.RGB
        elif (
            isinstance(color, tuple)
            and len(color) == 2
            and all(isinstance(v, int) for v in color)
        ):
            if color[0] > 9:
                raise ValueError("Theme color must be in the range 0-9.")
            if color[1] > 5:
                raise ValueError("Theme shade must be in the range 0-5.")

            self._theme_color = color
            self._type = ColorTypes.THEME
        else:
            raise ValueError(
                "Invalid color value. Must be a string, integer, or tuple."
            )

    def __repr__(self):
        """
        Return a string representation of the Color instance.
        """
        if self._type == ColorTypes.RGB:
            value = f"0x{self._rgb_value:06X}"
        else:
            value = f"Theme({self._theme_color[0]}, {self._theme_color[1]})"

        return (
            f"Color("
            f"value={value}, "
            f"type={self._type.name}, "
            f"is_automatic={self._is_automatic})"
        )

    @staticmethod
    def from_value(value: Union["Color", str]) -> "Color":
        """
        Convert a string to a Color instance or return the Color instance if
        already provided. This is mainly used for backward compatibility
        support.

        Args:
            value (Union[Color, str]): A Color instance or a string representing
            a color.

        Returns:
            Color: A Color instance.
        """
        if isinstance(value, Color):
            return value

        if isinstance(value, str):
            return Color(value)

        raise TypeError("Value must be a Color instance or a string.")

    @staticmethod
    def rgb(color: int) -> "Color":
        """
        Create a user-defined RGB color.

        Args:
            color (int): An RGB value in the range 0x000000 (black) to 0xFFFFFF (white).

        Returns:
            Color: A Color object representing the RGB color.
        """
        if color > 0xFFFFFF:
            raise ValueError("RGB color must be in the range 0x000000 - 0xFFFFFF.")
        return Color(color)

    @staticmethod
    def theme(color: int, shade: int) -> "Color":
        """
        Create a theme color.

        Args:
            color (int): The theme color index (0-9).
            shade (int): The theme shade index (0-5).

        Returns:
            Color: A Color object representing the theme color.
        """
        if color > 9:
            raise ValueError("Theme color must be in the range 0-9.")
        if shade > 5:
            raise ValueError("Theme shade must be in the range 0-5.")
        return Color((color, shade))

    def _parse_string_color(self, value: str):
        """
        Convert a hex string or named color to an RGB value.

        Returns:
            int: The RGB value.
        """
        # Named colors used in conjunction with various set_xxx_color methods to
        # convert a color name into an RGB value. These colors are for backward
        # compatibility with older versions of Excel.
        named_colors = {
            "red": 0xFF0000,
            "blue": 0x0000FF,
            "cyan": 0x00FFFF,
            "gray": 0x808080,
            "lime": 0x00FF00,
            "navy": 0x000080,
            "pink": 0xFF00FF,
            "black": 0x000000,
            "brown": 0x800000,
            "green": 0x008000,
            "white": 0xFFFFFF,
            "orange": 0xFF6600,
            "purple": 0x800080,
            "silver": 0xC0C0C0,
            "yellow": 0xFFFF00,
            "magenta": 0xFF00FF,
        }

        color = value.lstrip("#").lower()

        if color == "automatic":
            self._is_automatic = True
            self._rgb_value = 0x000000
        elif color in named_colors:
            self._rgb_value = named_colors[color]
        else:
            try:
                self._rgb_value = int(color, 16)
            except ValueError as e:
                raise ValueError(f"Invalid color value: {value}") from e

    def _rgb_hex_value(self) -> str:
        """
        Get the RGB hex value for the color.

        Returns:
            str: The RGB hex value as a string.
        """
        if self._is_automatic:
            # Default to black for automatic colors.
            return "000000"

        if self._type == ColorTypes.THEME:
            # Default to black for theme colors.
            return "000000"

        return f"{self._rgb_value:06X}"

    def _vml_rgb_hex_value(self) -> str:
        """
        Get the RGB hex value for a VML fill color in "#rrggbb" format.

        Returns:
            str: The RGB hex value as a string.
        """
        if self._is_automatic:
            # Default VML color for non-RGB colors.
            return "#ffffe1"

        return f"#{self._rgb_hex_value().lower()}"

    def _argb_hex_value(self) -> str:
        """
        Get the ARGB hex value for the color. The alpha channel is always FF.

        Returns:
            str: The ARGB hex value as a string.
        """
        return f"FF{self._rgb_hex_value()}"

    def _attributes(self):
        """
        Convert the color into a set of "rgb" or "theme/tint" attributes used in
        color-related Style XML elements.

        Returns:
            list[tuple[str, str]]: A list of key-value pairs representing the
            attributes.
        """
        # pylint: disable=too-many-return-statements
        # pylint: disable=no-else-return
        if self._type == ColorTypes.THEME:
            color, shade = self._theme_color

            # The first 3 columns of colors in the theme palette are different
            # from the others.
            if color == 0:
                if shade == 1:
                    return [("theme", str(color)), ("tint", "-4.9989318521683403E-2")]
                elif shade == 2:
                    return [("theme", str(color)), ("tint", "-0.14999847407452621")]
                elif shade == 3:
                    return [("theme", str(color)), ("tint", "-0.249977111117893")]
                elif shade == 4:
                    return [("theme", str(color)), ("tint", "-0.34998626667073579")]
                elif shade == 5:
                    return [("theme", str(color)), ("tint", "-0.499984740745262")]
                else:
                    return [("theme", str(color))]

            elif color == 1:
                if shade == 1:
                    return [("theme", str(color)), ("tint", "0.499984740745262")]
                elif shade == 2:
                    return [("theme", str(color)), ("tint", "0.34998626667073579")]
                elif shade == 3:
                    return [("theme", str(color)), ("tint", "0.249977111117893")]
                elif shade == 4:
                    return [("theme", str(color)), ("tint", "0.14999847407452621")]
                elif shade == 5:
                    return [("theme", str(color)), ("tint", "4.9989318521683403E-2")]
                else:
                    return [("theme", str(color))]

            elif color == 2:
                if shade == 1:
                    return [("theme", str(color)), ("tint", "-9.9978637043366805E-2")]
                elif shade == 2:
                    return [("theme", str(color)), ("tint", "-0.249977111117893")]
                elif shade == 3:
                    return [("theme", str(color)), ("tint", "-0.499984740745262")]
                elif shade == 4:
                    return [("theme", str(color)), ("tint", "-0.749992370372631")]
                elif shade == 5:
                    return [("theme", str(color)), ("tint", "-0.89999084444715716")]
                else:
                    return [("theme", str(color))]

            else:
                if shade == 1:
                    return [("theme", str(color)), ("tint", "0.79998168889431442")]
                elif shade == 2:
                    return [("theme", str(color)), ("tint", "0.59999389629810485")]
                elif shade == 3:
                    return [("theme", str(color)), ("tint", "0.39997558519241921")]
                elif shade == 4:
                    return [("theme", str(color)), ("tint", "-0.249977111117893")]
                elif shade == 5:
                    return [("theme", str(color)), ("tint", "-0.499984740745262")]
                else:
                    return [("theme", str(color))]

        # Handle RGB color.
        elif self._type == ColorTypes.RGB:
            return [("rgb", self._argb_hex_value())]

        # Default case for other colors.
        return []
