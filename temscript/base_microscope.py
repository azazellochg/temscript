import logging
import os.path
import platform

from .utils.constants import *
from .utils.enums import AcqImageFileFormat


class BaseMicroscope:

    def __init__(self, address=None, timeout=None, simulate=False, logLevel=logging.INFO):
        self._tem = None
        self._tem_adv = None

        logging.basicConfig(level=logLevel,
                            handlers=[
                                logging.FileHandler("debug.log"),
                                logging.StreamHandler()])

        if simulate:
            raise NotImplementedError()
        elif address is None:
            # local connection
            if platform.system() == "Windows":
                self._createInstrument()
            else:
                raise NotImplementedError("Running locally is only supported for Windows platform")
        else:
            raise NotImplementedError()

    def _createInstrument(self):
        """ Try to use both std and advanced scripting. """
        from comtypes.client import CreateObject
        try:
            self._tem_adv = CreateObject(SCRIPTING_ADV)
            logging.info("Connected to %s" % SCRIPTING_ADV)
        except:
            logging.info("Could not connect to %s" % SCRIPTING_ADV)
        try:
            self._tem = CreateObject(SCRIPTING_STD)
            logging.info("Connected to %s" % SCRIPTING_STD)
        except:
            self._tem = CreateObject(SCRIPTING_TECNAI)
            logging.info("Connected to %s" % SCRIPTING_TECNAI)


class Image:
    """ Acquired image object. """
    def __init__(self, obj, isAdvanced=False, **kwargs):
        self._img = obj
        self._isAdvanced = isAdvanced
        self.name = None if isAdvanced else obj.Name
        self.width = obj.Width
        self.height = obj.Height
        self.bit_depth = obj.BitDepth if isAdvanced else obj.Depth
        self.data = obj.AsSafeArray
        self.metadata = self._get_metadata(obj) if isAdvanced else None

    def _get_metadata(self, obj):
        return {item.Key: item.ValueAsString for item in obj.Metadata}

    def save(self, filename, normalize=False):
        """ Save acquired image to a file.

        :param filename: File path
        :type filename: str
        :param normalize: Normalize image
        :type normalize: bool
        """
        if self._isAdvanced:
            self._img.SaveToFile(filename, normalize=normalize)
        else:
            fmt = os.path.splitext(filename)[1].upper()
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
