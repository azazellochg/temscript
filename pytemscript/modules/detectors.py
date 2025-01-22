import logging
from ..utils.enums import AcqShutterMode, ScreenPosition


class Detectors:
    """ CCD/DDD, film/plate and STEM detectors. """
    def __init__(self, client):
        self._client = client
        self._has_film = None

        if self._client.has_advanced_iface:
            # CSA is supported by Ceta 1, Ceta 2, Falcon 3, Falcon 4
            self._tem_csa = self._client.get_from_cache("tem_adv.Acquisitions.CameraSingleAcquisition")
            if self._client.has("tem_adv.Acquisitions.CameraContinuousAcquisition"):
                # CCA is supported by Ceta 2
                self._tem_cca = self._client.get_from_cache("tem_adv.Acquisitions.CameraContinuousAcquisition")
                self._cca_cameras = [c.Name for c in self._tem_cca.SupportedCameras]
            else:
                self._tem_cca = None
                logging.info("Continuous acquisition not supported.")

    @property
    def __has_film(self):
        if self._has_film is None:
            self._has_film = self._client.has("tem.Camera.Stock")
        return self._has_film

    @property
    def cameras(self):
        """ Returns a dict with parameters for all cameras. """
        tem_cameras = dict()
        cameras = self._client.get_from_cache("tem.Acquisition.Cameras")
        for cam in cameras:
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
        if not self._client.has_advanced_iface:
            return tem_cameras

        for cam in self._tem_csa.SupportedCameras:
            self._tem_csa.Camera = cam
            param = self._tem_csa.CameraSettings.Capabilities
            tem_cameras[cam.Name] = {
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

            if self._tem_cca is not None and cam.Name in self._cca_cameras:
                self._tem_cca.Camera = cam
                param = self._tem_cca.CameraSettings.Capabilities
                tem_cameras[cam.Name].update({
                    "supports_recording": getattr(param, 'SupportsRecording', False)
                })

        return tem_cameras

    @property
    def stem_detectors(self):
        """ Returns a dict with STEM detectors parameters. """
        stem_detectors = dict()
        detectors = self._client.get_from_cache("tem.Acquisition.Detectors")
        for d in detectors:
            info = d.Info
            name = info.Name
            stem_detectors[name] = {
                "type": "STEM_DETECTOR",
                "binnings": [int(b) for b in info.Binnings]
            }
        return stem_detectors

    @property
    def screen(self):
        """ Fluorescent screen position. (read/write)"""
        return ScreenPosition(self._client.get("tem.Camera.MainScreen")).name

    @screen.setter
    def screen(self, value):
        self._client.set("tem.Camera.MainScreen", value)

    @property
    def film_settings(self):
        """ Returns a dict with film settings.
        Note: The plate camera has become obsolete with Win7 so
        most of the existing functions are no longer supported.
        """
        if self.__has_film:
            camera = self._client.get_from_cache("tem.Camera")
            return {
                "stock": camera.Stock,  # Int
                "exposure_time": camera.ManualExposureTime,
                "film_text": camera.FilmText,
                "exposure_number": camera.ExposureNumber,
                "user_code": camera.Usercode,  # 3 digits
                "screen_current": camera.ScreenCurrent * 1e9  # check if works without film
            }
        else:
            logging.error("No film/plate device detected.")
