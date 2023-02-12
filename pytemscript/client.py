from .base_microscope import BaseMicroscope
from .objects import *
from .utils.misc import rgetattr, rsetattr, rhasattr, rexecattr


class Microscope:
    """ Local client interface for the microscope.
    Creating an instance of this class will also create COM interfaces for the TEM.

    :param useLD: Connect to LowDose server on microscope PC (limited control only)
    :type useLD: bool
    :param useTecnaiCCD: Connect to TecnaiCCD plugin on microscope PC that controls Digital Micrograph (may be faster than via TIA / std scripting)
    :type useTecnaiCCD: bool
    :param useSEMCCD: Connect to SerialEMCCD plugin on Gatan PC that controls Digital Micrograph (may be faster than via TIA / std scripting)
    :type useSEMCCD: bool
    """
    def __init__(self, useLD=True, useTecnaiCCD=False, useSEMCCD=False):
        logging.basicConfig(level=logging.INFO,
                            datefmt='%d/%b/%Y %H:%M:%S',
                            format='[%(asctime)s] %(message)s',
                            handlers=[
                                logging.FileHandler("microscope.log", "w", "utf-8"),
                                logging.StreamHandler()])

        # Create all COM interfaces
        self._scope = BaseMicroscope(useLD, useTecnaiCCD, useSEMCCD)

        if useTecnaiCCD:
            if self._scope.tecnai_ccd is None:
                raise RuntimeError("Could not use Tecnai CCD plugin, "
                                   "please set useTecnaiCCD=False")
            else:
                from .plugins.tecnai_ccd_plugin import TecnaiCCDPlugin
                self._scope.tecnai_ccd_plugin = TecnaiCCDPlugin(self._scope)

        if useSEMCCD:
            if self._scope.sem_ccd is None:
                raise RuntimeError("Could not use SerialEM CCD plugin, "
                                   "please set useSEMCCD=False")
            else:
                from .plugins.serialem_ccd_plugin import SerialEMCCDPlugin
                self._sem_ccd_plugin = SerialEMCCDPlugin(self._scope)

        #self.acquisition = Acquisition(self._scope)
        #self.detectors = Detectors(self._scope)
        self.gun = Gun(self._scope)
        self.optics = Optics(self._scope)
        self.stem = Stem(self._scope)
        self.temperature = Temperature(self._scope)
        self.vacuum = Vacuum(self._scope)
        self.autoloader = Autoloader(self._scope)
        self.stage = Stage(self._scope)
        self.piezo_stage = PiezoStage(self._scope)
        #self.apertures = Apertures(self._scope)

        if self._scope.tem_adv is not None:
            self.user_door = UserDoor(self._scope)
            self.energy_filter = EnergyFilter(self._scope)

        if useLD:
            self.lowdose = LowDose(self._scope)

    def get(self, attrname):
        return rgetattr(self._scope, attrname)

    def exec(self, attrname, *args, **kwargs):
        attrname = attrname.rstrip("()")
        return rexecattr(self._scope, attrname, *args, **kwargs)

    def has(self, attrname):
        return rhasattr(self._scope, attrname)

    def set(self, attrname, value, vector=False, limits=None):
        if vector:
            values = list(map(float, value))
            if len(values) != 2:
                raise ValueError("Expected two values (X, Y) for Vector attribute %s" % attrname)

            if limits is not None:
                for v in values:
                    if not (limits[0] <= v <= limits[1]):
                        raise ValueError("%s is outside of range %s" % (v, limits))

            vector = rgetattr(self._scope, attrname)
            vector.X = values[0]
            vector.Y = values[1]
            rsetattr(self._scope, attrname, vector)
        else:
            rsetattr(self._scope, attrname, value)

    @property
    def family(self):
        """ Returns the microscope product family / platform. """
        return ProductFamily(self.get("tem.Configuration.ProductFamily")).name

    @property
    def condenser_system(self):
        """ Returns the type of condenser lens system: two or three lenses. """
        return CondenserLensSystem(self.get("tem.Configuration.CondenserLensSystem")).name

    @property
    def user_buttons(self):
        """ Returns a dict with assigned hand panels buttons. """
        return {b.Name: b.Label for b in self.get("tem.UserButtons")}
