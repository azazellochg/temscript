from .base_microscope import *
from .utils.properties import *


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
    def is_dewars_filling(self):
        return self.tem_temp_control.DewarsAreBusyFilling

    @property
    def dewars_remaining_time(self):
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
        if self.tem_autoloader.AutoLoaderAvailable:
            self.tem_autoloader.LoadCartridge(slot)

    def unload_cartridge(self):
        if self.tem_autoloader.AutoLoaderAvailable:
            self.tem_autoloader.UnloadCartridge()

    def run_inventory(self):
        if self.tem_autoloader.AutoLoaderAvailable:
            self.tem_autoloader.PerformCassetteInventory()

    def get_slot_status(self, slot):
        if self.tem_autoloader.AutoLoaderAvailable:
            status = self.tem_autoloader.SlotStatus(slot)
            return parse_enum(CassetteSlotStatus, status)

    @property
    def number_of_cassette_slots(self):
        if self.tem_autoloader.AutoLoaderAvailable:
            return self.tem_autoloader.NumberOfCassetteSlots


class Stage:

    def __init__(self, tem_stage):
        self.tem_stage = tem_stage


class Vacuum:
    """ Vacuum functions. """

    def __init__(self, microscope):
        self.tem_vacuum = microscope.tem.Vacuum

    @property
    def status(self):
        return self.tem_vacuum.Status

    @property
    def is_buffer_running(self):
        return self.tem_vacuum.PVPRunning

    @property
    def is_column_open(self):
        return self.tem_vacuum.ColumnValvesOpen

    def run_buffer_cycle(self):
        self.tem_vacuum.RunBufferCycle()

    @property
    def gauges(self):
        gauges = {}
        for g in self.tem_vacuum.Gauges:
            g.Read()
            status = GaugeStatus(g.Status)

            name = g.Name
            if status == GaugeStatus.UNDERFLOW:
                gauges[name] = "UNDERFLOW"
            elif status == GaugeStatus.OVERFLOW:
                gauges[name] = "OVERFLOW"
            elif status == GaugeStatus.VALID:
                gauges[name] = g.Pressure

        return gauges


class Optics(BaseOptics):
    """ Gun, Projection, Illumination, Apertures, VPP. """
    
    def __init__(self, microscope):
        super().__init__(microscope)
        self.mode = self.tem_control.InstrumentMode
        self.voltage = Voltage(self.tem_gun)

    @property
    def is_stem_available(self):
        return self.tem_control.StemAvailable

    @property
    def voltage_max(self):
        return self.tem_gun.HTMaxValue
