import os
import logging
from ..utils.enums import ImagePixelType, AcqImageFileFormat, StageAxes


class Vector:
    """ Utility object with two float attributes. """
    def __init__(self, x, y):
        self.x = x
        self.y = y
        self._min = None
        self._max = None

    def __repr__(self):
        return "Vector(x=%f, y=%f)" % (self.x, self.y)

    def set_limits(self, min_value, max_value):
        """Set the range limits for the vector."""
        self._min = min_value
        self._max = max_value

    @property
    def has_limits(self):
        """Check if range limits are set."""
        return self._min is not None and self._max is not None

    def check_limits(self):
        """Validate that the vector's values are within the set limits."""
        if self.has_limits:
            if any(v < self._min or v > self._max for v in self.components):
                msg = "One or more values (%s) are outside of range (%f, %f)" % (self.components, self.x, self.y)
                logging.error(msg)
                raise ValueError(msg)

    @property
    def components(self):
        """Return the vector components as a tuple."""
        return self.x, self.y


class StagePosition:
    """ Utility object for the stage position. """
    def __init__(self, **kwargs):
        self.coords = kwargs
        self.axes = 0

        for key, value in kwargs.items():
            if key not in 'xyzab':
                raise ValueError("Unexpected axis: %s" % key)
            self.axes |= getattr(StageAxes, key.upper())

    def __repr__(self):
        return "StagePosition(axes=%s, values=%s)" % (self.axes, self.coords)


class BaseImage:
    """ Acquired image basic object. """
    def __init__(self, obj, name=None, isAdvanced=False, **kwargs):
        self._img = obj
        self._name = name
        self._isAdvanced = isAdvanced
        self._kwargs = kwargs

    def _get_metadata(self, obj):
        raise NotImplementedError

    @property
    def name(self):
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

    def save(self, filename, normalize=False):
        """ Save acquired image to a file.

        :param filename: File path
        :type filename: str
        :param normalize: Normalize image, only for non-MRC format
        :type normalize: bool
        """
        raise NotImplementedError


class Image(BaseImage):
    """ Acquired image object. """
    def __init__(self, obj, name=None, isAdvanced=False, **kwargs):
        super().__init__(obj, name, isAdvanced, **kwargs)

    def _get_metadata(self, obj):
        return {item.Key: item.ValueAsString for item in obj.Metadata}

    @property
    def width(self):
        """ Image width in pixels. """
        return self._img.Width

    @property
    def height(self):
        """ Image height in pixels. """
        return self._img.Height

    @property
    def bit_depth(self):
        """ Bit depth. """
        return self._img.BitDepth if self._isAdvanced else self._img.Depth

    @property
    def pixel_type(self):
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

    def save(self, filename, normalize=False):
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
                    fmt = AcqImageFileFormat[fmt].value
                except KeyError:
                    raise NotImplementedError("Format %s is not supported" % fmt)
                self._img.AsFile(filename, fmt, normalize)
