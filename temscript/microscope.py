from .base_microscope import *


class Microscope(BaseMicroscope):
    """ Main class. """

    def __init__(self, address=None, timeout=None, simulate=False, logLevel=logging.INFO):
        super().__init__(address, timeout, simulate, logLevel)

        self.acquisition = Acquisition(self)
        self.detectors = Detectors(self)
        self.optics = Optics(self)
        self.temperature = Temperature(self)
        self.vacuum = Vacuum(self)
        self.sample = Sample(self)

    @property
    def family(self):
        return self.tem.Configuration.ProductFamily

    @property
    def user_buttons(self):
        buttons = {}
        for b in self.tem.UserButtons:
            buttons[b.Name] = b.Label
        return buttons


class Acquisition(BaseAcquisition):
    """ Image acquisition functions. """

    def acquire_tem_image(self, cameraName, **kwargs):
        self._acquire(cameraName, **kwargs)

    def acquire_stem_image(self, cameraName, **kwargs):
        self._acquire(cameraName, stem=True, **kwargs)


class Detectors(BaseDetectors):
    """ CCD/DDD, plate and STEM detectors. """

    def cameras(self):
        self.get_cameras()

    def stem_detectors(self):
        self.get_stem_detectors()


class Temperature:
    """ LN dewars and temperature controls. """

    def __init__(self, microscope):
        self.tem_temp_control = microscope.tem.TemperatureControl

    def force_refill(self):
        if self.tem_temp_control.TemperatureControlAvailable:
            self.tem_temp_control.ForceRefill()
        else:
            raise Exception("TemperatureControl is not available")

    def dewar_level(self, dewar):
        if self.tem_temp_control.TemperatureControlAvailable:
            dewar = parse_enum(RefrigerantDewar, dewar)
            return self.tem_temp_control.RefrigerantLevel(dewar)
        else:
            raise Exception("TemperatureControl is not available")

    @property
    def is_filling(self):
        return self.tem_temp_control.DewarsAreBusyFilling

    @property
    def remaining_time(self):
        return self.tem_temp_control.DewarsRemainingType


class Sample(BaseSample):
    """ Autoloader and Stage functions. """

    def __init__(self, microscope):
        super().__init__(microscope)
        self.autoloader = Autoloader(self.tem_autoloader)
        self.stage = Stage(self.tem_stage)


class Autoloader:

    def __init__(self, tem_autoloader):
        self.tem_autoloader = tem_autoloader

    def load_cartridge(self, slot):
        self.tem_autoloader.LoadCartridge(slot)

    def unload_cartridge(self):
        self.tem_autoloader.UnloadCartridge()

    def run_inventory(self):
        self.tem_autoloader.PerformCassetteInventory()

    def get_slot_status(self, slot):
        status = self.tem_autoloader.SlotStatus(slot)
        return parse_enum(CassetteSlotStatus, status)


class Stage:

    def __init__(self, tem_stage):
        self.tem_stage = tem_stage


class Vacuum:
    """ Vacuum functions. """

    def __init__(self, microscope):
        self.tem_vacuum = microscope.tem.Vacuum

    def is_column_open(self):
        return self.tem_vacuum.ColumnValvesOpen

    def run_buffer_cycle(self):
        self.tem_vacuum.RunBufferCycle()


class Optics(BaseOptics):
    """ Gun, Projection, Illumination, Apertures, VPP. """
    
    def __init__(self, microscope):
        super().__init__(microscope)
        self.mode = InstrumentMode(self.tem_control)
        self.voltage = Voltage(self.tem_gun)

    def is_stem_available(self):
        return self.tem_control.StemAvailable

    @property
    def voltage_max(self):
        return self.tem_gun.HTMaxValue

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
