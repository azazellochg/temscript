from typing import Optional, Dict
import os
import math
import logging
from pathlib import Path

from ..utils.enums import (ImagePixelType, AcqImageFileFormat, StageAxes,
                           GaugeStatus, GaugePressureLevel, AcqShutterMode,
                           MechanismId, MechanismState, MeasurementUnitType)


class Vector:
    """ Utility object with two float attributes. """
    def __init__(self, x: float, y: float):
        self.x = x
        self.y = y
        self._min: Optional[float] = None
        self._max: Optional[float] = None

    def __repr__(self):
        return "Vector(x=%f, y=%f)" % (self.x, self.y)

    def set_limits(self, min_value: float, max_value: float) -> None:
        """Set the range limits for the vector."""
        self._min = min_value
        self._max = max_value

    @property
    def has_limits(self) -> bool:
        """Check if range limits are set."""
        return self._min is not None and self._max is not None

    def check_limits(self) -> None:
        """Validate that the vector's values are within the set limits."""
        if self.has_limits:
            if any(v < self._min or v > self._max for v in self.components):
                msg = "One or more values (%s) are outside of range (%f, %f)" % (self.components, self.x, self.y)
                logging.error(msg)
                raise ValueError(msg)

    @property
    def components(self) -> tuple:
        """Return the vector components as a tuple."""
        return self.x, self.y


class BaseImage:
    """ Acquired image basic object. """
    def __init__(self,
                 obj,
                 name: Optional[str] = None,
                 isAdvanced: bool = False,
                 **kwargs):
        self._img = obj
        self._name = name
        self._isAdvanced = isAdvanced
        self._kwargs = kwargs

    def _get_metadata(self, obj):
        raise NotImplementedError

    @property
    def name(self) -> Optional[str]:
        """ Image name. """
        return self._name if self._isAdvanced else self._img.Name

    @property
    def width(self):
        """ Image width in pixels. """
        return None

    @property
    def height(self):
        """ Image height in pixels. """
        return None

    @property
    def bit_depth(self):
        """ Bit depth. """
        return None

    @property
    def pixel_type(self):
        """ Image pixels type: uint, int or float. """
        return None

    @property
    def data(self):
        """ Returns actual image object as numpy array. """
        return None

    @property
    def metadata(self):
        """ Returns a metadata dict for advanced camera image. """
        return self._get_metadata(self._img) if self._isAdvanced else None

    def save(self, filename: Path, normalize: bool = False):
        """ Save acquired image to a file.

        :param filename: File path
        :type filename: str
        :param normalize: Normalize image, only for non-MRC format
        :type normalize: bool
        """
        raise NotImplementedError


class Image(BaseImage):
    """ Acquired image object. """
    def __init__(self,
                 obj,
                 name: Optional[str] = None,
                 isAdvanced: bool = False,
                 **kwargs):
        super().__init__(obj, name, isAdvanced, **kwargs)

    def _get_metadata(self, obj) -> Dict:
        return {item.Key: item.ValueAsString for item in obj.Metadata}

    @property
    def width(self) -> int:
        """ Image width in pixels. """
        return self._img.Width

    @property
    def height(self) -> int:
        """ Image height in pixels. """
        return self._img.Height

    @property
    def bit_depth(self) -> str:
        """ Bit depth. """
        return self._img.BitDepth if self._isAdvanced else self._img.Depth

    @property
    def pixel_type(self) -> str:
        """ Image pixels type: uint, int or float. """
        if self._isAdvanced:
            return ImagePixelType(self._img.PixelType).name
        else:
            return ImagePixelType.SIGNED_INT.name

    @property
    def data(self):
        """ Returns actual image object as numpy int32 array. """
        from comtypes.safearray import safearray_as_ndarray
        with safearray_as_ndarray:
            data = self._img.AsSafeArray
        return data

    def save(self, filename: Path, normalize: bool = False) -> None:
        """ Save acquired image to a file.

        :param filename: File path
        :type filename: str
        :param normalize: Normalize image, only for non-MRC format
        :type normalize: bool
        """
        fmt = os.path.splitext(filename)[1].upper().replace(".", "")
        if fmt == "MRC":
            logging.info("Convert to int16 since MRC does not support int32")
            import mrcfile
            with mrcfile.new(filename) as mrc:
                if self.metadata is not None:
                    mrc.voxel_size = float(self.metadata['PixelSize.Width']) * 1e10
                mrc.set_data(self.data.astype("int16"))
        else:
            # use scripting method to save in other formats
            if self._isAdvanced:
                self._img.SaveToFile(filename)
            else:
                try:
                    fn_format = AcqImageFileFormat[fmt].value
                except KeyError:
                    raise NotImplementedError("Format %s is not supported" % fmt)
                self._img.AsFile(filename, fn_format, normalize)


class SpecialObj:
    """ Wrapper class for complex methods to be executed on a COM object. """
    def __init__(self, com_object, func: str = None, **kwargs):
        self.com_object = com_object
        self.func = func
        self.kwargs = kwargs

    def execute(self):
        method = getattr(self, self.func, None)
        if callable(method):
            return method(**self.kwargs)
        else:
            raise AttributeError("Function %s is not implemented for %s" % (
                self.func, SpecialObj.__name__))


class Detectors(SpecialObj):
    """ CCD/DDD, film/plate and STEM detectors."""

    def show_film_settings(self) -> Dict:
        """ Returns a dict with film settings. """
        camera = self.com_object
        return {
            "stock": camera.Stock,  # Int
            "exposure_time": camera.ManualExposureTime,
            "film_text": camera.FilmText,
            "exposure_number": camera.ExposureNumber,
            "user_code": camera.Usercode,  # 3 digits
            "screen_current": camera.ScreenCurrent * 1e9  # check if works without film
        }

    def show_stem_detectors(self) -> Dict:
        """ Returns a dict with STEM detectors parameters. """
        stem_detectors = dict()
        for d in self.com_object:
            info = d.Info
            name = info.Name
            stem_detectors[name] = {
                "type": "STEM_DETECTOR",
                "binnings": [int(b) for b in info.Binnings]
            }
        return stem_detectors

    def show_cameras(self) -> Dict:
        """ Returns a dict with parameters for all TEM cameras. """
        tem_cameras = dict()

        for cam in self.com_object:
            info = cam.Info
            param = cam.AcqParams
            name = info.Name
            tem_cameras[name] = {
                "type": "CAMERA",
                "height": info.Height,
                "width": info.Width,
                "pixel_size(um)": (info.PixelSize.X / 1e-6, info.PixelSize.Y / 1e-6),
                "binnings": [int(b) for b in info.Binnings],
                "shutter_modes": [AcqShutterMode(x).name for x in info.ShutterModes],
                "pre_exposure_limits(s)": (param.MinPreExposureTime, param.MaxPreExposureTime),
                "pre_exposure_pause_limits(s)": (param.MinPreExposurePauseTime,
                                                 param.MaxPreExposurePauseTime)
            }

        return tem_cameras

    def show_cameras_csa(self) -> Dict:
        """ Returns a dict with parameters for all TEM cameras that support CSA. """
        csa_cameras = dict()

        for cam in self.com_object.SupportedCameras:
            self.com_object.Camera = cam
            param = self.com_object.CameraSettings.Capabilities
            csa_cameras[cam.Name] = {
                "type": "CAMERA_ADVANCED",
                "height": cam.Height,
                "width": cam.Width,
                "pixel_size(um)": (cam.PixelSize.Width / 1e-6, cam.PixelSize.Height / 1e-6),
                "binnings": [int(b.Width) for b in param.SupportedBinnings],
                "exposure_time_range(s)": (param.ExposureTimeRange.Begin,
                                           param.ExposureTimeRange.End),
                "supports_dose_fractions": param.SupportsDoseFractions,
                "max_number_of_fractions": param.MaximumNumberOfDoseFractions,
                "supports_drift_correction": param.SupportsDriftCorrection,
                "supports_electron_counting": param.SupportsElectronCounting,
                "supports_eer": getattr(param, 'SupportsEER', False)
            }

        return csa_cameras

    def show_cameras_cca(self, tem_cameras: Dict) -> Dict:
        """ Update input dict with parameters for all TEM cameras that support CCA. """
        for cam in self.com_object.SupportedCameras:
            if cam.Name in tem_cameras:
                self.com_object.Camera = cam
                param = self.com_object.CameraSettings.Capabilities
                tem_cameras[cam.Name].update({
                    "supports_recording": getattr(param, 'SupportsRecording', False)
                })

        return tem_cameras


class StagePosition(SpecialObj):
    """ Stage functions. """

    def set(self,
            axes: int = 0,
            speed: Optional[float] = None,
            method: str = "MoveTo",
            **kwargs) -> None:
        """ Execute stage move to a new position. """
        if method not in ["MoveTo", "Goto", "GoToWithSpeed"]:
            raise NotImplementedError("Method %s is not implemented" % method)

        pos = self.com_object.Position
        for key, value in kwargs.items():
            setattr(pos, key.upper(), float(value))

        if speed is not None:
            getattr(self.com_object, method)(pos, axes, speed)
        else:
            getattr(self.com_object, method)(pos, axes)

    def get(self, a=False, b=False) -> Dict:
        """ The current position or speed of the stage/piezostage (x,y,z in um).
        Set a and b to True if you want to retrieve them as well.
        """
        pos = {key: getattr(self.com_object, key) * 1e6 for key in 'XYZ'}
        if a:
            pos['a'] = math.degrees(self.com_object.A)
            pos['b'] = None
        if b:
            pos['b'] = math.degrees(self.com_object.B)

        return pos

    def limits(self) -> Dict:
        """ Returns a dict with stage move limits. """
        limits = {}
        for axis in 'xyzab':
            data = self.com_object.AxisData(StageAxes[axis.upper()])
            limits[axis] = {
                'min': data.MinPos,
                'max': data.MaxPos,
                'unit': MeasurementUnitType(data.UnitType).name
            }

        return limits


class Gauges(SpecialObj):
    """ Vacuum gauges methods. """

    def show(self) -> Dict:
        """ Returns a dict with vacuum gauges information. """
        gauges = {}
        for g in self.com_object:
            # g.Read()
            if g.Status == GaugeStatus.UNDEFINED:
                # set manually if undefined, otherwise fails
                pressure_level = GaugePressureLevel.UNDEFINED.name
            else:
                pressure_level = GaugePressureLevel(g.PressureLevel).name

            gauges[g.Name] = {
                "status": GaugeStatus(g.Status).name,
                "pressure": g.Pressure,
                "trip_level": pressure_level
            }

        return gauges


class Apertures(SpecialObj):
    """ Apertures controls. """

    def show(self) -> Dict:
        """ Returns a dict with apertures information. """
        apertures = {}
        for ap in self.com_object:
            apertures[MechanismId(ap.Id).name] = {
                "retractable": ap.IsRetractable,
                "state": MechanismState(ap.State).name,
                "sizes": [a.Diameter for a in ap.ApertureCollection]
            }

        return apertures

    def _find_aperture(self, name: str):
        """ Helper method to find the aperture object by name. """
        name = name.upper()
        for ap in self.com_object:
            if name == MechanismId(ap.Id).name:
                return ap
        raise KeyError("No aperture with name %s" % name)

    def enable(self, name: str) -> None:
        ap = self._find_aperture(name)
        ap.Enable()

    def disable(self, name: str) -> None:
        ap = self._find_aperture(name)
        ap.Disable()

    def retract(self, name: str) -> None:
        ap = self._find_aperture(name)
        if ap.IsRetractable:
            ap.Retract()
        else:
            raise NotImplementedError("Aperture %s is not retractable" % name)

    def select(self, name: str, size: int) -> None:
        ap = self._find_aperture(name)
        if ap.State == MechanismState.DISABLED:
            ap.Enable()
        for a in ap.ApertureCollection:
            if a.Diameter == size:
                ap.SelectAperture(a)
                if ap.SelectedAperture.Diameter == size:
                    return
                else:
                    raise RuntimeError("Could not select aperture!")


class Buttons(SpecialObj):
    """ User buttons methods. """

    def show(self) -> Dict:
        """ Returns a dict with buttons assignment. """
        buttons = {}
        for b in self.com_object:
            buttons[b.Name] = b.Label

        return buttons
