from typing import Dict
import logging

from .extras import SpecialObj
from ..utils.enums import ScreenPosition, AcqShutterMode


class DetectorsObj(SpecialObj):
    """ Wrapper around cameras COM object."""

    def show_film_settings(self) -> Dict:
        """ Returns a dict with film settings. """
        film = self.com_object
        return {
            "stock": film.Stock,  # Int
            "exposure_time": film.ManualExposureTime,
            "film_text": film.FilmText,
            "exposure_number": film.ExposureNumber,
            "user_code": film.Usercode,  # 3 digits
            "screen_current": film.ScreenCurrent * 1e9  # check if works without film
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


class Detectors:
    """ CCD/DDD, film/plate and STEM detectors. """
    def __init__(self, client):
        self._client = client
        self._has_film = None
        self._has_cca = False

        # CCA is supported by Ceta 2
        if (self._client.has_advanced_iface and
                self._client.has("tem_adv.Acquisitions.CameraContinuousAcquisition")):
            self._has_cca = True
        else:
            logging.info("Continuous acquisition not supported.")

    @property
    def __has_film(self) -> bool:
        if self._has_film is None:
            self._has_film = self._client.has("tem.Camera.Stock")
        return self._has_film

    @property
    def cameras(self) -> Dict:
        """ Returns a dict with parameters for all TEM cameras. """
        tem_cameras = self._client.call("tem.Acquisition.Cameras", 
                                        obj=DetectorsObj,
                                        method="show_cameras")

        if not self._client.has_advanced_iface:
            return tem_cameras

        # CSA is supported by Ceta 1, Ceta 2, Falcon 3, Falcon 4
        csa_cameras = self._client.call("tem_adv.Acquisitions.CameraSingleAcquisition",
                                        obj=DetectorsObj, func="show_cameras_csa")
        tem_cameras.update(csa_cameras)

        # CCA is supported by Ceta 2
        if self._has_cca:
            tem_cameras = self._client.call("tem_adv.Acquisitions.CameraContinuousAcquisition",
                                            obj=DetectorsObj, func="show_cameras_cca",
                                            tem_cameras=tem_cameras)

        return tem_cameras

    @property
    def stem_detectors(self) -> Dict:
        """ Returns a dict with STEM detectors parameters. """
        return self._client.call("tem.Acquisition.Detectors",
                                 obj=DetectorsObj,
                                 func="show_stem_detectors")

    @property
    def screen(self) -> str:
        """ Fluorescent screen position. (read/write)"""
        return ScreenPosition(self._client.get("tem.Camera.MainScreen")).name

    @screen.setter
    def screen(self, value: ScreenPosition) -> None:
        self._client.set("tem.Camera.MainScreen", value)

    @property
    def film_settings(self) -> Dict:
        """ Returns a dict with film settings.
        Note: The plate camera has become obsolete with Win7 so
        most of the existing functions are no longer supported.
        """
        if self.__has_film:
            return self._client.call("tem.Camera", obj=DetectorsObj,
                                     func="show_film_settings")
        else:
            logging.error("No film/plate device detected.")
            return {}
