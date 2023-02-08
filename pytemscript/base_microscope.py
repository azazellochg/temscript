import logging
import platform
import sys
import atexit

from .utils.constants import *
from .utils.enums import TEMScriptingError


class BaseMicroscope:
    """ Base class that handles COM interface connections. """
    def __init__(self, useLD, useTecnaiCCD, useSEMCCD, logLevel=logging.INFO, remote=False):
        self._tem = None
        self._tem_adv = None
        self._lowdose = None
        self._tecnai_ccd = None
        self._sem_ccd = None

        if not remote:
            logging.basicConfig(level=logLevel,
                                datefmt='%d/%b/%Y %H:%M:%S',
                                format='[%(asctime)s] %(message)s',
                                handlers=[
                                    logging.FileHandler("debug.log", "w", "utf-8"),
                                    logging.StreamHandler()])

        if platform.system() == "Windows":
            self._initialize(useLD, useTecnaiCCD, useSEMCCD)
            atexit.register(self._close)
        else:
            raise NotImplementedError("Running locally is only supported for Windows platform")

    def _createCOMObject(self, progId):
        """ Connect to a COM interface. """
        try:
            import comtypes.client
            obj = comtypes.client.CreateObject(progId)
            logging.info("Connected to %s" % progId)
            return obj
        except:
            logging.info("Could not connect to %s" % progId)
            return None

    def _initialize(self, useLD, useTecnaiCCD, useSEMCCD):
        """ Wrapper to create interfaces as requested. """
        import comtypes
        try:
            comtypes.CoInitializeEx(comtypes.COINIT_MULTITHREADED)
        except OSError:
            comtypes.CoInitialize()

        self._tem_adv = self._createCOMObject(SCRIPTING_ADV)
        self._tem = self._createCOMObject(SCRIPTING_STD)

        if self._tem is None:  # try Tecnai instead
            self._tem = self._createCOMObject(SCRIPTING_TECNAI)

        if useLD:
            self._lowdose = self._createCOMObject(SCRIPTING_LOWDOSE)
        if useTecnaiCCD:
            self._tecnai_ccd = self._createCOMObject(SCRIPTING_TECNAI_CCD)
            if self._tecnai_ccd is None:
                self._tecnai_ccd = self._createCOMObject(SCRIPTING_TECNAI_CCD2)
            import comtypes.gen.TECNAICCDLib
        if useSEMCCD:
            from .utils.gatan_socket import SocketFuncs
            self._sem_ccd = SocketFuncs()

    @staticmethod
    def handle_com_error(com_error):
        """ Try catching COM error. """
        try:
            default = TEMScriptingError.E_NOT_OK.value
            err = TEMScriptingError(int(getattr(com_error, 'hresult', default))).name
            logging.info('COM error: %s' % err)
        except ValueError:
            logging.info('Exception : %s' % sys.exc_info()[1])

    def _close(self):
        import comtypes
        comtypes.CoUninitialize()


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


class Vector:
    """ Vector object/property. """

    @staticmethod
    def set(obj, attr_name, values, range=None):
        values = list(map(float, values))
        if len(values) != 2:
            raise ValueError("Expected two values for Vector attribute %s" % attr_name)

        if range is not None:
            for v in values:
                if not(range[0] <= v <= range[1]):
                    raise ValueError("%s is outside of range %s" % (v, range))

        vector = getattr(obj, attr_name)
        vector.X = values[0]
        vector.Y = values[1]
        setattr(obj, attr_name, vector)
