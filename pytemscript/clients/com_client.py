import logging
from ..utils.misc import rgetattr, rsetattr
from ..base_microscope import BaseMicroscope


class COMClient:
    """ Local COM client interface for the microscope.
    Creating an instance of this class will also create COM interfaces for the TEM.

    :param useLD: Connect to LowDose server on microscope PC (limited control only)
    :type useLD: bool
    :param useTecnaiCCD: Connect to TecnaiCCD plugin on microscope PC that controls Digital Micrograph (maybe faster than via TIA / std scripting)
    :type useTecnaiCCD: bool
    """
    def __init__(self, useLD=False, useTecnaiCCD=False):
        logging.basicConfig(level=logging.INFO,
                            datefmt='%d/%b/%Y %H:%M:%S',
                            format='[%(asctime)s] %(message)s',
                            handlers=[
                                logging.FileHandler("com_client.log", "w", "utf-8"),
                                logging.StreamHandler()])

        # Create all COM interfaces
        self._scope = BaseMicroscope(useLD, useTecnaiCCD)

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
        return self.cache.get(attrname)

    def clear_cache(self, attrname):
        if attrname in self.cache:
            del self.cache[attrname]

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
