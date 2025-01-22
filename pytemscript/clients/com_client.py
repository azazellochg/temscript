import logging
import platform
import sys
sys.coinit_flags = 0
import comtypes  # COM is initialized automatically for the thread that imports this module for the first time
import comtypes.client
from ..utils.misc import rgetattr, rsetattr
from ..utils.constants import *
from ..utils.enums import TEMScriptingError


class COMBase:
    """ Base class that handles COM interface connections. """
    def __init__(self, useLD=False, useTecnaiCCD=False):
        self.tem = None
        self.tem_adv = None
        self.tem_lowdose = None
        self.tecnai_ccd = None

        if platform.system() == "Windows":
            logging.getLogger("comtypes").setLevel(logging.INFO)
            self._initialize(useLD, useTecnaiCCD)
        else:
            raise NotImplementedError("Running locally is only supported for Windows platform")

    @staticmethod
    def _createCOMObject(progId):
        """ Connect to a COM interface. """
        try:
            obj = comtypes.client.CreateObject(progId)
            logging.info("Connected to %s" % progId)
            return obj
        except:
            logging.info("Could not connect to %s" % progId)
            return None

    def _initialize(self, useLD, useTecnaiCCD):
        """ Wrapper to create interfaces as requested. """
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


class COMClient:
    """ Local COM client interface for the microscope.
    Creating an instance of this class will also create COM interfaces for the TEM.

    :param useLD: Connect to LowDose server on microscope PC (limited control only)
    :type useLD: bool
    :param useTecnaiCCD: Connect to TecnaiCCD plugin on microscope PC that controls Digital Micrograph (maybe faster than via TIA / std scripting)
    :type useTecnaiCCD: bool
    :param debug: Print debug messages
    :type debug: bool
    """
    def __init__(self, useLD=False, useTecnaiCCD=False, debug=False):
        logging.basicConfig(level=logging.DEBUG if debug else logging.INFO,
                            datefmt='%d/%b/%Y %H:%M:%S',
                            format='[%(asctime)s] %(message)s',
                            handlers=[
                                logging.FileHandler("com_client.log", "w", "utf-8"),
                                logging.StreamHandler()])

        # Create all COM interfaces
        self._scope = COMBase(useLD, useTecnaiCCD)

        if useTecnaiCCD:
            if self._scope.tecnai_ccd is None:
                raise RuntimeError("Could not use Tecnai CCD plugin, "
                                   "please set useTecnaiCCD=False")
            else:
                from ..plugins.tecnai_ccd_plugin import TecnaiCCDPlugin
                self._ccd_plugin = TecnaiCCDPlugin(self._scope.tecnai_ccd)

        self.cache = {}

    @property
    def has_advanced_iface(self):
        return self._scope.tem_adv is not None

    @property
    def has_lowdose_iface(self):
        return self._scope.tem_lowdose is not None

    @property
    def has_ccd_iface(self):
        return self._scope.tecnai_ccd is not None

    def get(self, attrname):
        return rgetattr(self._scope, attrname)

    def get_from_cache(self, attrname):
        if attrname not in self.cache:
            self.cache[attrname] = self.get(attrname)
        return self.cache[attrname]

    def clear_cache(self, attrname):
        if attrname in self.cache:
            del self.cache[attrname]

    def has(self, attrname):
        """ GET request with cache support. Should be used only for attributes
        that do not change over the session. """
        if attrname not in self.cache:
            try:
                _ = self.get(attrname)
                self.cache[attrname] = True
            except AttributeError:
                self.cache[attrname] = False

        return self.cache[attrname]

    def call(self, attrname, *args, **kwargs):
        attrname = attrname.rstrip("()")
        return rgetattr(self._scope, attrname, *args, **kwargs)

    def set(self, attrname, value, **kwargs):
        if kwargs.get("vector"):
            values = list(map(float, value))
            if len(values) != 2:
                msg = "Expected two values (X, Y) for Vector %s" % attrname
                logging.error(msg)
                raise ValueError(msg)

            limits = kwargs.get("limits")
            if limits and any(v < limits[0] or v > limits[1] for v in values):
                msg = "One or more values (%s) are outside of range (%s)" % (values, limits)
                logging.error(msg)
                raise ValueError(msg)

            vector = self.get(attrname)
            vector.X = values[0]
            vector.Y = values[1]
            rsetattr(self._scope, attrname, vector)
        else:
            rsetattr(self._scope, attrname, value)
