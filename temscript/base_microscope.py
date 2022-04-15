import logging
import os.path
from comtypes.client import CreateObject

from utils.constants import *
from utils.enums import AcqImageFileFormat


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
            self._createInstrument()
        else:
            raise NotImplementedError()

    def _createInstrument(self):
        """ Try to use both std and advanced scripting. """
        try:
            self._tem_adv = CreateObject(SCRIPTING_ADV)
            self._tem = CreateObject(SCRIPTING_STD)
            logging.info(f"Connected to {SCRIPTING_ADV} and {SCRIPTING_STD}")
        except:
            logging.info(f"Could not connect to {SCRIPTING_ADV}")
            self._tem = CreateObject(SCRIPTING_STD)
        else:
            self._tem = CreateObject(SCRIPTING_TECNAI)
            logging.info(f"Connected to {SCRIPTING_TECNAI}")
        finally:
            raise Exception("Could not connect to the instrument")

    def _check_licensing(self):
        try:
            #self._lic_adv = CreateObject(LICENSE_ADV)
            self._lic_cam = CreateObject(LICENSE_ADV_CAM)
        except:
            logging.info(f"Could not connect to advanced instrument")


class Image:
    """ Acquired image object. """
    def __init__(self, obj, isAdvanced=False, **kwargs):
        self._img = obj
        self._isAdvanced = isAdvanced

        self.name = None if isAdvanced else obj.Name
        self.width = obj.Width
        self.height = obj.Height
        self.bit_depth = obj.BitDepth if isAdvanced else obj.Depth
        self.data = obj._AsSafeArray
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
                fmt = AcqImageFileFormat[fmt]
            except KeyError:
                raise NotImplementedError(f"Format {fmt} is not supported.")
            self._img.SaveToFile(filename, fmt, normalize)


class Vector:
    """ Vector object/property. """
    def __init__(self, com_obj, name='', readonly=False):
        self._com_obj = com_obj
        self._name = name
        self._readonly = readonly

    def __get__(self, obj, objtype=None):
        result = getattr(self._com_obj, self._name)
        return result.X, result.Y

    def __set__(self, obj, value):
        if self._readonly:
            raise AttributeError(f"Attribute {self._name} is not writable")
        value = [float(c) for c in value]
        if len(value) != 2:
            raise ValueError(f"Expected two items for attribute {self._name}")

        vector = getattr(self._com_obj, self._name)
        vector.X = value[0]
        vector.Y = value[1]

        setattr(self._com_obj, self._name, vector)
