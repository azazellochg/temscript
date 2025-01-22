import os
import logging
from ..utils.enums import ImagePixelType, AcqImageFileFormat


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
