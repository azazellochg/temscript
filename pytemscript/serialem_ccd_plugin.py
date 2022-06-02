import logging
import os
import time

from .utils.enums import *
from .base_microscope import BaseImage


class SerialEMCCDPlugin:
    """ Main class that uses SerialEM CCD plugin on Gatan PC. """
    def __init__(self, microscope):
        if microscope._sem_ccd is not None:
            self._socket = microscope._sem_ccd
            self._img_params = dict()

    def _find_camera(self, name):
        """Find camera index by name. """
        for i in range(self._socket.GetNumberOfCameras()):
            if self._socket.CameraName == name:
                return i
        raise KeyError("No camera with name %s" % name)

    def acquire_image(self, cameraName, size=AcqImageSize.FULL, exp_time=1, binning=1, camerasize=1024, **kwargs):
        raise NotImplementedError

    def _set_camera_param(self, name, size, exp_time, binning, camerasize, **kwargs):
        """ Find the TEM camera and set its params. """
        raise NotImplementedError

    def _run_command(self, command, *args):
        raise NotImplementedError


class Image(BaseImage):
    """ Acquired image object. """
    def __init__(self, obj, name=None, **kwargs):
        super().__init__(obj, name, isAdvanced=False, **kwargs)
