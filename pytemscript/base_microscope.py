import logging
import platform
import sys
import atexit

from utils.constants import *
from utils.enums import TEMScriptingError


class BaseMicroscope:
    """ Base class that handles COM interface connections. """
    def __init__(self, useLD=False, useTecnaiCCD=False):
        self.tem = None
        self.tem_adv = None
        self.tem_lowdose = None
        self.tecnai_ccd = None

        if platform.system() == "Windows":
            self._initialize(useLD, useTecnaiCCD)
            atexit.register(self._close)
        else:
            raise NotImplementedError("Running locally is only supported for Windows platform")

    @staticmethod
    def _createCOMObject(progId):
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
            logging.error('COM error: %s' % err)
        except ValueError:
            logging.error('Exception : %s' % sys.exc_info()[1])

    @staticmethod
    def _close():
        import comtypes
        comtypes.CoUninitialize()
