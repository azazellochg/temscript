import logging
from comtypes.client import CreateObject

from utils.constants import *
from utils.methods import *
from utils.enums import *


class BaseMicroscope:

    def __init__(self, address=None, timeout=None, simulate=False, logLevel=logging.INFO):
        self.tem = None
        self.tem_adv = None

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
        finally:
            raise Exception("Could not connect to the instrument")


class BaseAcquisition:

    def __init__(self, microscope):
        self.tem = microscope.tem
        self.tem_adv = microscope.tem_adv
        self.tem_acq = self.tem.Acquisition
        self.tem_csa = self.tem_adv.Acquisitions.CameraSingleAcquisition
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


class BaseDetectors:

    def __init__(self, microscope):
        self.tem = microscope.tem
        self.tem_adv = microscope.tem_adv
        self.tem_acq = self.tem.Acquisition
        self.tem_csa = self.tem_adv.Acquisitions.CameraSingleAcquisition
        self.cameras = {}
        self.stem_detectors = {}

    def get_cameras(self):
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

    def get_stem_detectors(self):
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

    def get_screen_position(self):
        return self.tem.Camera.MainScreen.name

    def set_screen_position(self, mode):
        mode = parse_enum(ScreenPosition, mode)
        self.tem.Camera.MainScreen = mode


class BaseSample:

    def __init__(self, microscope):
        self.tem = microscope.tem
        self.tem_autoloader = self.tem.AutoLoader
        self.tem_stage = self.tem.Stage


class BaseOptics:

    def __init__(self, microscope):
        self.tem = microscope.tem
        self.tem_adv = microscope.tem_adv
        self.tem_gun = self.tem.Gun
        self.tem_gun1 = None
        self.tem_illumination = self.tem.Illumination
        self.tem_projection = self.tem.Projection
        self.tem_control = self.tem.InstrumentModeControl
        self.tem_vpp = self.tem_adv.Phaseplate


### properties

class InstrumentMode:
    def __init__(self, tem_control):
        self.tem_control = tem_control

    def __get__(self, obj, objType=None):
        return self.tem_control.InstrumentMode.name

    def __set__(self, obj, value):
        mode = parse_enum(InstrumentMode, value)
        self.tem_control.InstrumentMode = mode


class Voltage:
    def __init__(self, tem_gun):
        self.tem_gun = tem_gun

    def __get__(self, obj, objType=None):
        state = self.tem_gun.HTState
        if state == HighTensionState.ON:
            return self.tem_gun.HTValue * 1e-3
        else:
            return 0.0

    def __set__(self, obj, value):
        self.tem_gun.HTValue = value

