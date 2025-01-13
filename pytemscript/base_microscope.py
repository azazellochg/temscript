import logging
import platform
import sys
import atexit

from .utils.constants import *
from .utils.enums import TEMScriptingError


class BaseMicroscope:
    """ Base class that handles COM interface connections. """
    def __init__(self, useLD=False, useTecnaiCCD=False):
        self.tem = None
        self.tem_adv = None
        self.tem_lowdose = None
        self.tecnai_ccd = None
        self.sem_ccd = None

        if platform.system() == "Windows":
            self._initialize(useLD, useTecnaiCCD)
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

    def _initialize(self, useLD, useTecnaiCCD):
        """ Wrapper to create interfaces as requested. """
        import comtypes
        try:
            comtypes.CoInitializeEx(comtypes.COINIT_MULTITHREADED)
        except OSError:
            comtypes.CoInitialize()

        self.tem_adv = self._createCOMObject(SCRIPTING_ADV)
        self.tem = self._createCOMObject(SCRIPTING_STD)

        if self.tem is None:  # try Tecnai instead
            self.tem = self._createCOMObject(SCRIPTING_TECNAI)

        if useLD:
            self.tem_lowdose = self._createCOMObject(SCRIPTING_LOWDOSE)
        if useTecnaiCCD:
            self.tecnai_ccd = self._createCOMObject(SCRIPTING_TECNAI_CCD)
            if self.tecnai_ccd is None:
                self.tecnai_ccd = self._createCOMObject(SCRIPTING_TECNAI_CCD2)
            import comtypes.gen.TECNAICCDLib

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
