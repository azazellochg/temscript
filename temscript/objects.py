from .base_microscope import set_enum_attr_from_dict, set_attr_from_dict
from .utils.enums import *


class Acquisition:

    def __init__(self, microscope):
        self.tem = microscope.tem
        self.tem_adv = microscope.tem_adv
        self.acq = self.tem.Acquisition
        self.csa = self.tem_adv.Acquisitions.CameraSingleAcquisition
        self.isAdvanced = False

    def _find_camera(self, name):
        """Find camera object by name. Check adv scripting first. """
        for cam in self.csa.SupportedCameras:
            if cam.Name == name:
                self.isAdvanced = True
                return cam
        for cam in self.acq.Cameras:
            if cam.Info.Name == name:
                return cam
        raise KeyError("No camera with name %s" % name)

    def _find_stem_detector(self, name):
        """Find STEM detector object by name"""
        for stem in self.acq.Detectors:
            if stem.Info.Name == name:
                return stem
        raise KeyError("No STEM detector with name %s" % name)

    def _check_binning(self, binning):
        """ Check if input binning is in SupportedBinnings for a single-acquisition camera.
        Assume that the camera is set: self._tem_csa.Camera = camera.
        Returns Binning object.
        """
        param = self.csa.CameraSettings.Capabilities
        for b in param.SupportedBinnings:
            if int(b.Width) == int(binning):
                return b
        return False

    def set_camera_param(self, name, params, ignore_errors=False):
        camera = self._find_camera(name)

        if self.isAdvanced:
            self.csa.Camera = camera
            settings = self.csa.CameraSettings
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

    def set_stem_detector_param(self, name, params, ignore_errors=False):
        det = self._find_stem_detector(name)
        info = det.Info
        set_attr_from_dict(info, 'Brightness', params, 'brightness')
        set_attr_from_dict(info, 'Contrast', params, 'contrast')

        param = self.acq.StemAcqParams
        set_enum_attr_from_dict(param, 'ImageSize', AcqImageSize, params, 'image_size', ignore_errors=ignore_errors)
        set_attr_from_dict(param, 'Binning', params, 'binning', ignore_errors=ignore_errors)
        set_attr_from_dict(param, 'DwellTime', params, 'dwell_time(s)', ignore_errors=ignore_errors)

        if not ignore_errors and params:
            raise ValueError("Unknown keys in parameter dictionary.")

    def acquire_tem_image(self, cameraName, **kwargs):
        result = {}
        self.set_camera_param(cameraName, **kwargs)
        if self.isAdvanced:
            pass
        else:
            self.acq.RemoveAllAcqDevices()
            self.acq.AddAcqDeviceByName(cameraName)
            img = self.acq.AcquireImages()
            result[img.Name] = img._AsSafeArray
            return result

    def acquire_stem_image(self, cameraName, **kwargs):
        pass


class Detectors:

    def __init__(self, microscope):
        self.tem = microscope.tem
        self.tem_adv = microscope.tem_adv
        self.acq = self.tem.Acquisition
        self.csa = self.tem_adv.Acquisitions.CameraSingleAcquisition
        self.cameras = microscope.cameras


class Optics:

    def __init__(self, microscope):
        self.tem = microscope.tem
        self.tem_adv = microscope.tem_adv


class Temperature:

    def __init__(self, microscope):
        self.tem = microscope.tem
        self.tem_adv = microscope.tem_adv


class Vacuum:

    def __init__(self, microscope):
        self.tem = microscope.tem
        self.tem_adv = microscope.tem_adv


class Sample:

    def __init__(self, microscope):
        self.tem = microscope.tem
        self.tem_adv = microscope.tem_adv
