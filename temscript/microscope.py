from comtypes.client import CreateObject
from http.client import HTTPConnection
import logging

#from .base_microscope import BaseMicroscope
from .objects import *
from utils.constants import *


class Microscope:
    """ Main class """

    def __init__(self, address=None, timeout=None, simulate=False, logLevel=logging.INFO):
        self.tem = None
        self.tem_adv = None
        self.cameras = dict()
        self.address = address
        self.timeout = timeout

        self.acquisition = Acquisition(self)
        self.detectors = Detectors(self)  # Cameras, STEM detectors
        self.optics = Optics(self)  # Gun, Projection, Illumination, Apertures, VPP
        self.temperature = Temperature(self)
        self.vacuum = Vacuum(self)
        self.sample = Sample(self)  # Autoloader and Stage

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
            self.tem_adv = CreateObject(SCRIPTING_ADV)
            self.tem = CreateObject(SCRIPTING_STD)
            logging.info(f"Connected to {SCRIPTING_ADV} and {SCRIPTING_STD}")
        except:
            logging.info(f"Could not connect to {SCRIPTING_ADV}")
            self.tem = CreateObject(SCRIPTING_STD)
        else:
            self.tem = CreateObject(SCRIPTING_TECNAI)
            logging.info(f"Connected to {SCRIPTING_TECNAI}")


'''

class Prop:
    """ property """
    def __get__(self, obj, objType=None):
        print(f"Getting {obj}")
        value = obj._age
        return value

    def __set__(self, obj, value):
        print(f"Setting {obj} to {value}")
        self.validate(value)
        obj._age = value

    def validate(self, value):
        if not isinstance(value, str):
            raise TypeError("Value must be a string")


class Acquisition:
    property1 = Prop()

    def do_smth(self):
        print(f"Got param from Big instance: {Big.get_param}")
        print(f"Executed do_smth")


class Big:
    acquisition = Acquisition()

    def __init__(self):
        self.param = None

    def run(self):
        self.param = os.path.dirname(__file__)

    @property
    def get_param(self):
        return getattr(Big, 'param', None)


big = Big()
big.run()

big.acquisition.property1 = "obj value"
print("Result:", big.acquisition.property1)
big.acquisition.do_smth()
'''
