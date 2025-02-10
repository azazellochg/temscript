from typing import Optional, Dict
import os
import math
import logging
from pathlib import Path

from ..utils.enums import (ImagePixelType, AcqImageFileFormat, StageAxes,
                           MeasurementUnitType)


class Vector:
    """ Utility object with two float attributes. """
    def __init__(self, x: float, y: float):
        self.x = x
        self.y = y
        self._min: Optional[float] = None
        self._max: Optional[float] = None

    def __repr__(self):
        return "Vector(x=%f, y=%f)" % (self.x, self.y)

    def set_limits(self, min_value: float, max_value: float) -> None:
        """Set the range limits for the vector."""
        self._min = min_value
        self._max = max_value

    @property
    def has_limits(self) -> bool:
        """Check if range limits are set."""
        return self._min is not None and self._max is not None

    def check_limits(self) -> None:
        """Validate that the vector's values are within the set limits."""
        if self.has_limits:
            if any(v < self._min or v > self._max for v in self.components):
                msg = "One or more values (%s) are outside of range (%f, %f)" % (self.components, self.x, self.y)
                logging.error(msg)
                raise ValueError(msg)

    @property
    def components(self) -> tuple:
        """Return the vector components as a tuple."""
        return self.x, self.y


class BaseImage:
    """ Acquired image basic object. """
    def __init__(self,
                 obj,
                 name: Optional[str] = None,
                 isAdvanced: bool = False,
                 **kwargs):
        self._img = obj
        self._name = name
        self._isAdvanced = isAdvanced
        self._kwargs = kwargs

    def _get_metadata(self, obj):
        raise NotImplementedError

    @property
    def name(self) -> Optional[str]:
        """ Image name. """
        return self._name if self._isAdvanced else self._img.Name

    @property
    def width(self):
        """ Image width in pixels. """
        return None

    @property
    def height(self):
        """ Image height in pixels. """
        return None

    @property
    def bit_depth(self):
        """ Bit depth. """
        return None

    @property
    def pixel_type(self):
        """ Image pixels type: uint, int or float. """
        return None

    @property
    def data(self):
        """ Returns actual image object as numpy array. """
        return None

    @property
    def metadata(self):
        """ Returns a metadata dict for advanced camera image. """
        return self._get_metadata(self._img) if self._isAdvanced else None

    def save(self, filename: Path, normalize: bool = False):
        """ Save acquired image to a file.

        :param filename: File path
        :type filename: str
        :param normalize: Normalize image, only for non-MRC format
        :type normalize: bool
        """
        raise NotImplementedError


class Image(BaseImage):
    """ Acquired image object. """
    def __init__(self,
                 obj,
                 name: Optional[str] = None,
                 isAdvanced: bool = False,
                 **kwargs):
        super().__init__(obj, name, isAdvanced, **kwargs)

    def _get_metadata(self, obj) -> Dict:
        return {item.Key: item.ValueAsString for item in obj.Metadata}

    @property
    def width(self) -> int:
        """ Image width in pixels. """
        return self._img.Width

    @property
    def height(self) -> int:
        """ Image height in pixels. """
        return self._img.Height

    @property
    def bit_depth(self) -> str:
        """ Bit depth. """
        return self._img.BitDepth if self._isAdvanced else self._img.Depth

    @property
    def pixel_type(self) -> str:
        """ Image pixels type: uint, int or float. """
        if self._isAdvanced:
            return ImagePixelType(self._img.PixelType).name
        else:
            return ImagePixelType.SIGNED_INT.name

    @property
    def data(self):
        """ Returns actual image object as numpy int32 array. """
        from comtypes.safearray import safearray_as_ndarray
        with safearray_as_ndarray:
            data = self._img.AsSafeArray
        return data

    def save(self, filename: Path, normalize: bool = False) -> None:
        """ Save acquired image to a file.

        :param filename: File path
        :type filename: str
        :param normalize: Normalize image, only for non-MRC format
        :type normalize: bool
        """
        fmt = os.path.splitext(filename)[1].upper().replace(".", "")
        if fmt == "MRC":
            logging.info("Convert to int16 since MRC does not support int32")
            import mrcfile
            with mrcfile.new(filename) as mrc:
                if self.metadata is not None:
                    mrc.voxel_size = float(self.metadata['PixelSize.Width']) * 1e10
                mrc.set_data(self.data.astype("int16"))
        else:
            # use scripting method to save in other formats
            if self._isAdvanced:
                self._img.SaveToFile(filename)
            else:
                try:
                    fn_format = AcqImageFileFormat[fmt].value
                except KeyError:
                    raise NotImplementedError("Format %s is not supported" % fmt)
                self._img.AsFile(filename, fn_format, normalize)


class SpecialObj:
    """ Wrapper class for complex methods to be executed on a COM object. """
    def __init__(self, com_object, func: str, **kwargs):
        self.com_object = com_object
        self.func = func
        self.kwargs = kwargs

    def execute(self):
        method = getattr(self, self.func)
        if callable(method):
            return method(**self.kwargs)
        else:
            raise AttributeError("Function %s is not implemented for %s" % (
                self.func, SpecialObj.__name__))


class StagePosition(SpecialObj):
    """ Wrapper around stage / piezo stage COM object. """

    def set(self,
            axes: int = 0,
            speed: Optional[float] = None,
            method: str = "MoveTo",
            **kwargs) -> None:
        """ Execute stage move to a new position. """
        if method not in ["MoveTo", "Goto", "GoToWithSpeed"]:
            raise NotImplementedError("Method %s is not implemented" % method)

        pos = self.com_object.Position
        for key, value in kwargs.items():
            setattr(pos, key.upper(), float(value))

        if speed is not None:
            getattr(self.com_object, method)(pos, axes, speed)
        else:
            getattr(self.com_object, method)(pos, axes)

    def get(self, a=False, b=False) -> Dict:
        """ The current position or speed of the stage/piezo stage (x,y,z in um).
        Set a and b to True if you want to retrieve them as well.
        """
        pos = {key: getattr(self.com_object, key) * 1e6 for key in 'XYZ'}
        if a:
            pos['a'] = math.degrees(self.com_object.A)
            pos['b'] = None
        if b:
            pos['b'] = math.degrees(self.com_object.B)

        return pos

    def limits(self) -> Dict:
        """ Returns a dict with stage move limits. """
        limits = {}
        for axis in 'xyzab':
            data = self.com_object.AxisData(StageAxes[axis.upper()])
            limits[axis] = {
                'min': data.MinPos,
                'max': data.MaxPos,
                'unit': MeasurementUnitType(data.UnitType).name
            }

        return limits
