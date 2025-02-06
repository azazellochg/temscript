from typing import Dict
import logging

from .extras import Detectors as DetectorsObj
from ..utils.enums import ScreenPosition


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
        tem_cameras = self._client.call("tem.Acquisition.Cameras", obj=DetectorsObj,
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
        return self._client.call("tem.Acquisition.Detectors", obj=DetectorsObj,
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
