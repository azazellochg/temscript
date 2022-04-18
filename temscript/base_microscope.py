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


class BaseProperty:
    __slots__ = 'com_object', 'name', 'readonly'

    def __init__(self, com_object, name=None, readonly=False):
        self.com_object = com_object
        self.name = name
        self.readonly = readonly


class VectorProperty(BaseProperty):
    """ Wrapper for the Vector COM scripting object. Stores two floats: X and Y.

    :param com_object: input COM object
    :param attr_name: attribute name for the COM object
    :type attr_name: str
    :param range: tuple with (min, max) values
    :type range: tuple
    :param readonly: whether the attribute is read only
    :type readonly: bool
    """
    __slots__ = 'com_object', 'name', 'range', 'readonly'

    def __init__(self, com_object, attr_name=None, range=None, readonly=False):
        super().__init__(com_object, attr_name, readonly=readonly)
        self.range = range

    def __get__(self, instance):
        if self.name is None:
            return self
        else:
            result = getattr(self.com_object, self.name)
            return [result.X, result.Y]

    def __set__(self, instance, values):
        if self.readonly:
            raise AttributeError("Attribute %s is not writable" % self.name)

        values = list(map(float, values))
        if len(values) != 2:
            raise ValueError("Expected two items for attribute" % self.name)

        if range is not None:
            err = "%s are outside of range %s" % (values, self.range)
            assert self.range[0] <= values[0] <= self.range[1], err
            assert self.range[0] <= values[1] <= self.range[1], err

        # New attr value should be set as a whole, since get creates a copy
        vector = getattr(self.com_object, self.name)
        vector.X, vector.Y = values[0], values[1]
        setattr(self.com_object, self.name, vector)


class EnumProperty(BaseProperty):
    """ Wrapper for the Enumeration COM scripting object. Stores IntEnum.

    :param com_object: input COM object
    :param attr_name: attribute name for the COM object
    :type attr_name: str
    :param enum_type: IntEnum class from enums.py
    :type enum_type: IntEnum
    :param readonly: whether the attribute is read only
    :type readonly: bool
    """
    __slots__ = 'com_object', 'enum_type', 'name', 'readonly'

    def __init__(self, com_object, attr_name=None, enum_type=None, readonly=False):
        super().__init__(com_object, attr_name, readonly=readonly)
        self.enum_type = enum_type

    def __get__(self, instance):
        if self.name is None:
            return self
        else:
            value = getattr(self.com_object, self.name)
            return self.enum_type(value).name

    def __set__(self, instance, value):
        if self.readonly:
            raise AttributeError("Attribute %s is not writable" % self.name)
        setattr(self.com_object, self.name, int(value))


class NumProperty(BaseProperty):
    """ Wrapper for the long/int property. Stores float/int value.

    :param com_object: input COM object
    :param attr_name: attribute name for the COM object
    :type attr_name: str
    :param attr_type: attribute value type, float or int
    :param readonly: whether the attribute is read only
    :type readonly: bool
    """
    __slots__ = 'com_object', 'attr_type', 'name', 'readonly'

    def __init__(self, com_object, attr_name=None, attr_type=None, readonly=False):
        super().__init__(com_object, attr_name, readonly=readonly)
        self.attr_type = attr_type

    def __get__(self, instance):
        if self.name is None:
            return self
        else:
            return getattr(self.com_object, self.name)

    def __set__(self, instance, value):
        if self.readonly:
            raise AttributeError("Attribute %s is not writable" % self.name)
        setattr(self.com_object, self.name, self.attr_type(value))
