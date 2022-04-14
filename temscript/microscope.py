from .base_microscope import BaseMicroscope
from .utils.properties import *
from .utils.methods import *
from .utils.enums import *


class Microscope(BaseMicroscope):
    """ Main class. """

    def __init__(self, address=None, timeout=None, simulate=False):
        super().__init__(address, timeout, simulate)

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
    def condenser_system(self):
        return self.tem.Configuration.CondernserLensSystem

    @property
    def user_buttons(self):
        buttons = {}
        for b in self.tem.UserButtons:
            buttons[b.Name] = b.Label
        return buttons


class Acquisition:
    """ Image acquisition functions. """

    def __init__(self, microscope):
        self.tem = microscope.tem
        self.tem_adv = microscope.tem_adv
        self.tem_acq = self.tem.Acquisition
        self.tem_csa = self.tem_adv.Acquisitions.CameraSingleAcquisition
        self.tem_cam = microscope.Camera
        self.isAdvanced = False

    def _find_camera(self, name):
        """Find camera object by name. Check adv scripting first. """
        for cam in self.tem_csa.SupportedCameras:
            if cam.Name == name:
                self.isAdvanced = True
                return cam
        for cam in self.tem_acq.Cameras:
            if cam.Info.Name == name:
                return cam
        raise KeyError("No camera with name %s" % name)

    def _find_stem_detector(self, name):
        """Find STEM detector object by name"""
        for stem in self.tem_acq.Detectors:
            if stem.Info.Name == name:
                return stem
        raise KeyError("No STEM detector with name %s" % name)

    def _check_binning(self, binning):
        """ Check if input binning is in SupportedBinnings for a single-acquisition camera.
        Assume that the camera is set: self._tem_csa.Camera = camera.
        Returns Binning object.
        """
        param = self.tem_csa.CameraSettings.Capabilities
        for b in param.SupportedBinnings:
            if int(b.Width) == int(binning):
                return b
        return False

    def _set_camera_param(self, name, params, ignore_errors=False):
        camera = self._find_camera(name)

        if self.isAdvanced:
            self.tem_csa.Camera = camera
            settings = self.tem_csa.CameraSettings
            set_enum_attr_from_dict(settings, 'ReadoutArea', AcqImageSize, params, 'image_size',
                                    ignore_errors=ignore_errors)

            binning = self._check_binning(params['binning'])
            if binning:
                params['binning'] = binning
                set_attr_from_dict(settings, 'Binning', params, 'binning', ignore_errors=ignore_errors)

            # Set exposure after binning, since it adjusted automatically when binning is set
            set_attr_from_dict(settings, 'ExposureTime', params, 'exposure(s)', ignore_errors=ignore_errors)

        else:
            info = camera.Info
            settings = camera.AcqParams
            set_enum_attr_from_dict(settings, 'ImageSize', AcqImageSize, params, 'image_size', ignore_errors=ignore_errors)
            set_attr_from_dict(settings, 'Binning', params, 'binning', ignore_errors=ignore_errors)
            set_enum_attr_from_dict(settings, 'ImageCorrection', AcqImageCorrection, params, 'correction',
                                    ignore_errors=ignore_errors)
            set_enum_attr_from_dict(settings, 'ExposureMode', AcqExposureMode, params, 'exposure_mode',
                                    ignore_errors=ignore_errors)
            set_enum_attr_from_dict(info, 'ShutterMode', AcqShutterMode, params, 'shutter_mode',
                                    ignore_errors=ignore_errors)
            set_attr_from_dict(settings, 'PreExposureTime', params, 'pre_exposure(s)', ignore_errors=ignore_errors)
            set_attr_from_dict(settings, 'PreExposurePauseTime', params, 'pre_exposure_pause(s)',
                               ignore_errors=ignore_errors)

            # Set exposure after binning, since it adjusted automatically when binning is set
            set_attr_from_dict(settings, 'ExposureTime', params, 'exposure(s)', ignore_errors=ignore_errors)

        if not ignore_errors and params:
            raise ValueError("Unknown keys in parameter dictionary.")

    def _set_stem_detector_param(self, name, params, ignore_errors=False):
        det = self._find_stem_detector(name)
        info = det.Info
        set_attr_from_dict(info, 'Brightness', params, 'brightness')
        set_attr_from_dict(info, 'Contrast', params, 'contrast')

        settings = self.tem_acq.StemAcqParams
        set_enum_attr_from_dict(settings, 'ImageSize', AcqImageSize, params, 'image_size', ignore_errors=ignore_errors)
        set_attr_from_dict(settings, 'Binning', params, 'binning', ignore_errors=ignore_errors)
        set_attr_from_dict(settings, 'DwellTime', params, 'dwell_time(s)', ignore_errors=ignore_errors)

        if not ignore_errors and params:
            raise ValueError("Unknown keys in parameter dictionary.")

    def _set_film_param(self, params, ignore_errors=False):
        set_attr_from_dict(self.tem_cam, 'FilmText', params, 'film_text',
                           ignore_errors=ignore_errors)
        set_attr_from_dict(self.tem_cam, 'ManualExposureTime', params, 'exposure_time(s)',
                           ignore_errors=ignore_errors)

    def _acquire(self, cameraName, stem=False, **kwargs):
        result = {}
        if stem:
            self._set_stem_detector_param(cameraName, **kwargs)
        else:
            self._set_camera_param(cameraName, **kwargs)
            if self.isAdvanced:
                img = self.tem_csa.Acquire()
                md = img.Metadata
                img_name = md['DetectorName'] + '_' + md['TimeStamp']
                result[img_name] = img._AsSafeArray
                return result

        self.tem_acq.RemoveAllAcqDevices()
        self.tem_acq.AddAcqDeviceByName(cameraName)
        img = self.tem_acq.AcquireImages()
        result[img.Name] = img._AsSafeArray

        return result

    def acquire_tem_image(self, cameraName, **kwargs):
        self._acquire(cameraName, **kwargs)

    def acquire_stem_image(self, cameraName, **kwargs):
        self._acquire(cameraName, stem=True, **kwargs)

    def acquire_film(self, **kwargs):
        if self.tem_cam.Stock > 0:
            self.tem_cam.PlateLabelDataType = PlateLabelDateFormat.DDMMYY
            self.tem_cam.ExposureNumber += 1
            self.tem_cam.MainScreen = ScreenPosition.UP
            self.tem_cam.ScreenDim = True
            self._set_film_param(**kwargs)
            self.tem_cam.TakeExposure()
        else:
            raise Exception("Plate stock is empty!")


class Detectors:
    """ CCD/DDD, plate and STEM detectors. """

    def __init__(self, microscope):
        self.tem = microscope.tem
        self.tem_adv = microscope.tem_adv
        self.tem_acq = self.tem.Acquisition
        self.tem_csa = self.tem_adv.Acquisitions.CameraSingleAcquisition
        self.tem_cam = microscope.Camera

        self.screen = EnumProperty(self.tem.Camera, ScreenPosition, 'MainScreen')

    def cameras(self):
        self.cameras = dict()
        for cam in self.tem_acq.Cameras:
            info = cam.Info
            param = cam.AcqParams
            name = info.Name
            self.cameras[name] = {
                "type": "CAMERA",
                "height": info.Height,
                "width": info.Width,
                "pixel_size(um)": tuple(size / 1e-6 for size in info.PixelSize),
                "binnings": [int(b) for b in info.Binnings],
                "shutter_modes": [AcqShutterMode(x).name for x in info.ShutterModes],
                "pre_exposure_limits(s)": (param.MinPreExposureTime, param.MaxPreExposureTime),
                "pre_exposure_pause_limits(s)": (param.MinPreExposurePauseTime, param.MaxPreExposurePauseTime)
            }
        for cam in self.tem_csa.SupportedCameras:
            self.tem_csa.Camera = cam
            param = self.tem_csa.CameraSettings.Capabilities
            self.cameras[cam.Name] = {
                "type": "CAMERA_ADVANCED",
                "height": cam.Height,
                "width": cam.Width,
                "pixel_size(um)": (cam.PixelSize.Width / 1e-6, cam.PixelSize.Height / 1e-6),
                "binnings": [int(b.Width) for b in param.SupportedBinnings],
                "exposure_time_range(s)": (param.ExposureTimeRange.Begin, param.ExposureTimeRange.End),
                "supports_dose_fractions": param.SupportsDoseFractions,
                "supports_drift_correction": param.SupportsDriftCorrection,
                "supports_electron_counting": param.SupportsElectronCounting,
                "supports_eer": param.SupportsEER
            }

        return self.cameras

    def stem_detectors(self):
        self.stem_detectors = dict()
        for det in self.tem_acq.Detectors:
            info = det.Info
            name = info.Name
            self.stem_detectors[name] = {
                "type": "STEM_DETECTOR",
                "binnings": [int(b) for b in info.Binnings],
                "brightness": info.Brightness,
                "contrast": info.Contrast
            }
        return self.stem_detectors

    @property
    def film_settings(self):
        return {
            "stock": self.tem_cam.Stock,
            "exposure_time(s)": self.tem_cam.ManualExposureTime,
            "film_text": self.tem_cam.FilmText,
            "exposure_number": self.tem_cam.ExposureNumber,
            "user_code": self.tem_cam.Usercode,
            "screen_current": self.tem_cam.ScreenCurrent
        }


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
        return self.tem_temp_control.DewarsRemainingTime


class Sample:
    """ Autoloader and Stage functions. """

    def __init__(self, microscope):
        self.tem = microscope.tem
        self.autoloader = Autoloader(self.tem.AutoLoader)
        self.stage = Stage(self.tem.Stage)


class Autoloader:
    """ Sample loading functions. """

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
            return CassetteSlotStatus(status).name

    @property
    def number_of_cassette_slots(self):
        if self.tem_autoloader.AutoLoaderAvailable:
            return self.tem_autoloader.NumberOfCassetteSlots


class Stage:
    """ Stage functions. """

    def __init__(self, tem_stage):
        self.tem_stage = tem_stage
        self.status = EnumProperty(self.tem_stage, StageStatus, "Status", readonly=True)
        self.holder = EnumProperty(self.tem_stage, StageHolderType, "Holder", readonly=True)

    def from_dict(self, position, values):
        axes = 0
        for key, value in values.items():
            if key not in 'xyzab':
                raise ValueError("Unexpected axes: %s" % key)
            attr_name = key.upper()
            setattr(position, attr_name, float(value))
            axes |= getattr(StageAxes, attr_name)
        return axes

    @property
    def position(self):
        pos = self.tem_stage.Position
        axes = 'xyzab'
        return {key: getattr(pos, key.upper()) for key in axes}

    def go_to(self, speed=None, **kwargs):
        pos = self.tem_stage.Position
        axes = self.from_dict(pos, kwargs)
        if speed:
            self.tem_stage.GoToWithSpeed(axes, speed)
        else:
            self.tem_stage.GoTo(axes)

    def move_to(self, **kwargs):
        pos = self.tem_stage.Position
        axes = self.from_dict(pos, kwargs)
        self.tem_stage.MoveTo(axes)

    @property
    def limits(self):
        result = dict()
        for axis in 'xyzab':
            data = self.tem_stage.AxisData(StageAxes[axis.upper()])
            result[axis] = {
                'min': data.MinPos,
                'max': data.MaxPos,
                'unit': MeasurementUnitType(data.UnitType).name
            }
        return result


class Vacuum:
    """ Vacuum functions. """

    def __init__(self, microscope):
        self.tem_vacuum = microscope.tem.Vacuum
        self.status = EnumProperty(self.tem_vacuum, VacuumStatus, "Status", readonly=True)

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
            gauges[g.Name] = {
                "status": GaugeStatus(g.Status).name,
                "pressure": g.Pressure,
                "level": GaugePressureLevel(g.PressureLevel).name
            }
        return gauges


class Optics:
    """ Gun, Projection, Illumination, Apertures, VPP. """
    
    def __init__(self, microscope):
        self.tem = microscope.tem
        self.tem_adv = microscope.tem_adv
        self.tem_gun = self.tem.Gun
        self.tem_gun1 = self.tem.Gun
        self.tem_illumination = self.tem.Illumination
        self.tem_projection = self.tem.Projection
        self.tem_control = self.tem.InstrumentModeControl
        self.tem_vpp = self.tem_adv.Phaseplate

        self.mode = EnumProperty(self.tem_control, InstrumentMode, "InstrumentMode")
        self.voltage_offset = self.tem_gun1.HighVoltageOffset
        self.shutter_override = self.tem.BlankerShutter.ShutterOverrideOn
        self.ht_state = EnumProperty(self.tem_gun, HighTensionState, "HTState")

    @property
    def is_stem_available(self):
        return self.tem_control.StemAvailable

    @property
    def voltage(self):
        state = self.tem_gun.HTState
        if state == HighTensionState.ON:
            return self.tem_gun.HTValue * 1e-3
        else:
            return 0.0

    @voltage.setter
    def voltage(self, value):
        self.tem_gun.HTValue = value

    @property
    def voltage_max(self):
        return self.tem_gun.HTMaxValue

    @property
    def voltage_offset_range(self):
        return self.tem_gun1.GetHighVoltageOffsetRange()

    @property
    def is_shutter_override(self):
        return self.tem.BlankerShutter
