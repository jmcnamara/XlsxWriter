###############################################################################
#
# Model3D - A class for representing 3D model objects in Excel.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2025, John McNamara, jmcnamara@cpan.org
#

import hashlib
import json
import os
import struct
from io import BytesIO
from pathlib import Path
from typing import Dict, Optional, Tuple, Union

from xlsxwriter.exceptions import UnsupportedImageFormat


class Model3D:
    """
    A class to represent a 3D model in an Excel worksheet.

    Supports GLB (binary glTF) format which is required by Excel.
    """

    def __init__(self, source: Union[str, Path, BytesIO]) -> None:
        """
        Initialize a Model3D instance.

        Args:
            source (Union[str, Path, BytesIO]): The filename, Path or BytesIO
            object of the 3D model (GLB format).
        """
        if isinstance(source, (str, Path)):
            self.filename = str(source)
            self.model_data: Optional[BytesIO] = None
            self.model_name = os.path.basename(source)
        elif isinstance(source, BytesIO):
            self.filename = ""
            self.model_data = source
            self.model_name = ""
        else:
            raise ValueError("Source must be a filename (str) or a BytesIO object.")

        self._row: int = 0
        self._col: int = 0
        self._x_offset: int = 0
        self._y_offset: int = 0
        self._x_scale: float = 1.0
        self._y_scale: float = 1.0
        self._z_scale: float = 1.0
        self._anchor: int = 2
        self._description: Optional[str] = None
        self._decorative: bool = False

        # 3D model specific properties
        self._model_extension: str = "glb"
        self._width: float = 200.0  # Default display width in pixels
        self._height: float = 200.0  # Default display height in pixels
        self._digest: Optional[str] = None

        # Camera and lighting defaults (matching Excel's defaults)
        self._camera_pos = (0, 0, 54040559)  # Camera position
        self._camera_up = (0, 36000000, 0)  # Camera up vector
        self._camera_look_at = (0, 0, 0)  # Look at point
        self._camera_fov = 2700000  # Field of view

        # Model transform defaults
        self._meter_per_unit = (1000000, 1000000)  # n/d ratio
        self._pre_trans = (0, 0, 0)  # Pre-translation
        self._scale = (1000000, 1000000, 1000000)  # Scale factors (n/d where d=1000000)
        self._post_trans = (0, 0, 0)  # Post-translation

        # Preview image (for fallback in older Excel versions)
        self._preview_image: Optional[bytes] = None

        self._get_model_properties()

    def __repr__(self) -> str:
        """
        Return a string representation of the main properties of the Model3D
        instance.
        """
        return (
            f"Model3D:\n"
            f"    filename   = {self.filename!r}\n"
            f"    model_name = {self.model_name!r}\n"
            f"    model_type = {self.model_type!r}\n"
            f"    width      = {self._width}\n"
            f"    height     = {self._height}\n"
        )

    @property
    def model_type(self) -> str:
        """Get the model type (e.g., 'GLB')."""
        return self._model_extension.upper()

    @property
    def width(self) -> float:
        """Get the display width of the model."""
        return self._width

    @property
    def height(self) -> float:
        """Get the display height of the model."""
        return self._height

    @property
    def description(self) -> Optional[str]:
        """Get the description/alt-text of the model."""
        return self._description

    @description.setter
    def description(self, value: str) -> None:
        """Set the description/alt-text of the model."""
        if value:
            self._description = value

    @property
    def decorative(self) -> bool:
        """Get whether the model is decorative."""
        return self._decorative

    @decorative.setter
    def decorative(self, value: bool) -> None:
        """Set whether the model is decorative."""
        self._decorative = value

    def _set_user_options(self, options: Optional[Dict] = None) -> None:
        """
        Handle the additional optional parameters to ``insert_3d_model()``.
        """
        if options is None:
            return

        self._anchor = options.get("object_position", self._anchor)
        self._x_scale = options.get("x_scale", self._x_scale)
        self._y_scale = options.get("y_scale", self._y_scale)
        self._z_scale = options.get("z_scale", self._z_scale)
        self._x_offset = options.get("x_offset", self._x_offset)
        self._y_offset = options.get("y_offset", self._y_offset)
        self._decorative = options.get("decorative", self._decorative)
        self._description = options.get("description", self._description)
        self._width = options.get("width", self._width)
        self._height = options.get("height", self._height)

        # Preview image for fallback
        self._preview_image = options.get("preview_image", self._preview_image)

    def _get_model_properties(self) -> None:
        """Extract properties from the 3D model file."""
        if self.model_data:
            data = self.model_data.getvalue()
        else:
            with open(self.filename, "rb") as fh:
                data = fh.read()

        # Get the model digest to check for duplicates
        self._digest = hashlib.sha256(data).hexdigest()

        # Validate GLB format
        if len(data) < 12:
            raise UnsupportedImageFormat(
                f"{self.filename}: File too small to be a valid GLB."
            )

        # GLB header: magic (4 bytes) + version (4 bytes) + length (4 bytes)
        magic = struct.unpack("<I", data[0:4])[0]
        if magic != 0x46546C67:  # 'glTF' in little-endian
            raise UnsupportedImageFormat(
                f"{self.filename}: Not a valid GLB file (invalid magic number)."
            )

        version = struct.unpack("<I", data[4:8])[0]
        if version != 2:
            raise UnsupportedImageFormat(
                f"{self.filename}: Unsupported GLB version {version}. Only version 2 is supported."
            )

        self._model_extension = "glb"

        # Try to extract bounds from the GLB to set reasonable defaults
        bounds = self._extract_glb_bounds(data)
        if bounds:
            # Use bounds to calculate aspect ratio for default display size
            min_bounds, max_bounds = bounds
            model_width = max_bounds[0] - min_bounds[0]
            model_height = max_bounds[1] - min_bounds[1]
            model_depth = max_bounds[2] - min_bounds[2]

            # Calculate pre-translation to center the model
            center_x = (min_bounds[0] + max_bounds[0]) / 2
            center_y = (min_bounds[1] + max_bounds[1]) / 2
            center_z = (min_bounds[2] + max_bounds[2]) / 2

            # Store as EMU-like units (multiply by 36000000 for Excel's coordinate system)
            scale = 36000000
            self._pre_trans = (
                int(-center_x * scale),
                int(-center_y * scale),
                int(-center_z * scale),
            )

            # Set aspect ratio for display
            if model_width > 0 and model_height > 0:
                aspect = model_width / model_height
                if aspect > 1:
                    self._width = 200.0
                    self._height = 200.0 / aspect
                else:
                    self._height = 200.0
                    self._width = 200.0 * aspect

    def _extract_glb_bounds(
        self, data: bytes
    ) -> Optional[Tuple[Tuple[float, float, float], Tuple[float, float, float]]]:
        """
        Extract bounding box from GLB file if available.

        Returns:
            Tuple of (min_bounds, max_bounds) or None if not found.
        """
        try:
            # GLB structure: header (12 bytes) + chunks
            offset = 12

            while offset < len(data):
                chunk_length = struct.unpack("<I", data[offset : offset + 4])[0]
                chunk_type = struct.unpack("<I", data[offset + 4 : offset + 8])[0]

                # JSON chunk type is 0x4E4F534A ('JSON' in little-endian)
                if chunk_type == 0x4E4F534A:
                    json_data = data[offset + 8 : offset + 8 + chunk_length]
                    gltf = json.loads(json_data.decode("utf-8"))

                    # Look for accessors with min/max bounds
                    if "accessors" in gltf:
                        for accessor in gltf["accessors"]:
                            if (
                                accessor.get("type") == "VEC3"
                                and "min" in accessor
                                and "max" in accessor
                            ):
                                min_bounds = tuple(accessor["min"])
                                max_bounds = tuple(accessor["max"])
                                return (min_bounds, max_bounds)

                    # If no accessor bounds, check for scene bounds in extras
                    if "scenes" in gltf and gltf["scenes"]:
                        scene = gltf["scenes"][0]
                        if "extras" in scene and "bounds" in scene["extras"]:
                            bounds = scene["extras"]["bounds"]
                            return (tuple(bounds["min"]), tuple(bounds["max"]))

                    break

                offset += 8 + chunk_length

        except (struct.error, json.JSONDecodeError, KeyError, IndexError):
            pass

        return None
