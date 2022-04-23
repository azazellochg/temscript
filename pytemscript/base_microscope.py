import logging
import os.path
import platform
import warnings

try:
    import comtypes.client
except ImportError:
    warnings.warn("Importing comtypes failed. Non-Windows platform?", ImportWarning)

from .utils.constants import *
from .utils.enums import AcqImageFileFormat, ImagePixelType


class BaseMicroscope:

    def __init__(self, address=None, timeout=None, simulate=False, logLevel=logging.INFO):
        self._tem = None
        self._tem_adv = None
        self._lowdose = None
        self._semccd = None
        self._fei_gatan = None
        self._address = address

        logging.basicConfig(level=logLevel,
                            handlers=[
                                logging.FileHandler("debug.log"),
                                logging.StreamHandler()])

        if simulate:
            raise NotImplementedError()
        elif self._address is None:
            if platform.system() == "Windows":
                self._createLocalInstrument()
            else:
                raise NotImplementedError("Running locally is only supported for Windows platform")
        else:
            raise NotImplementedError()

    def _createLocalInstrument(self):
        """ Try to use both std and advanced scripting. """
        try:
            comtypes.CoInitializeEx(comtypes.COINIT_MULTITHREADED)
        except WindowsError:
            comtypes.CoInitialize()
        try:
            self._tem_adv = comtypes.client.CreateObject(SCRIPTING_ADV)
            logging.info("Connected to %s" % SCRIPTING_ADV)
        except:
            logging.info("Could not connect to %s" % SCRIPTING_ADV)
        try:
            self._tem = comtypes.client.CreateObject(SCRIPTING_STD)
            logging.info("Connected to %s" % SCRIPTING_STD)
        except:
            self._tem = comtypes.client.CreateObject(SCRIPTING_TECNAI)
            logging.info("Connected to %s" % SCRIPTING_TECNAI)
        try:
            self._lowdose = comtypes.client.CreateObject(SCRIPTING_LOWDOSE)
            logging.info("Connected to %s" % SCRIPTING_LOWDOSE)
        except:
            logging.info("Could not connect to %s" % SCRIPTING_LOWDOSE)
        try:
            self._semccd = comtypes.client.CreateObject(SCRIPTING_SEM_CCD)
            logging.info("Connected to %s" % SCRIPTING_SEM_CCD)
        except:
            logging.info("Could not connect to %s" % SCRIPTING_SEM_CCD)
        try:
            self._fei_gatan = comtypes.client.CreateObject(SCRIPTING_FEI_GATAN_REMOTING)
            logging.info("Connected to %s" % SCRIPTING_FEI_GATAN_REMOTING)
        except:
            logging.info("Could not connect to %s" % SCRIPTING_FEI_GATAN_REMOTING)

    def __del__(self):
        if self._address is None:
            comtypes.CoUninitialize()


class Image:
    """ Acquired image object. """
    def __init__(self, obj, name=None, isAdvanced=False, **kwargs):
        self._img = obj
        self._name = name
        self._isAdvanced = isAdvanced

    def _get_metadata(self, obj):
        return {item.Key: item.ValueAsString for item in obj.Metadata}

    @property
    def name(self):
        """ Image name. """
        return self._name if self._isAdvanced else self._img.Name

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
        fmt = os.path.splitext(filename)[1].upper().replace(".", "")
        if fmt == "MRC":
            print("Convert to int16 since MRC does not support int32")
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


class Vector:
    """ Vector object/property. """

    @staticmethod
    def set(obj, attr_name, value, range=None):
        value = [float(c) for c in value]
        if len(value) != 2:
            raise ValueError("Expected two items for attribute %s" % attr_name)

        if range is not None:
            for v in value:
                if not(range[0] <= v <= range[1]):
                    raise ValueError("%s is outside of range %s" % (value, range))

        vector = getattr(obj, attr_name)
        vector.X = value[0]
        vector.Y = value[1]
        setattr(obj, attr_name, vector)
