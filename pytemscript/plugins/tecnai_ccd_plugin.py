import logging
import os
import time

from ..utils.enums import AcqImageSize, AcqMode, AcqSpeed, ImagePixelType
from ..modules.image import BaseImage


class TecnaiCCDPlugin:
    """ Main class that uses FEI Tecnai CCD plugin on microscope PC. """
    def __init__(self, com_iface):
            self._plugin = com_iface
            self._img_params = dict()

    def _find_camera(self, name):
        """Find camera index by name. """
        for i in range(self._plugin.NumberOfCameras):
            if self._plugin.CameraName == name:
                return i
        raise KeyError("No camera with name %s" % name)

    def acquire_image(self, cameraName, size=AcqImageSize.FULL, exp_time=1,
                      binning=1, camerasize=1024, **kwargs):
        self._set_camera_param(cameraName, size, exp_time, binning, camerasize, **kwargs)
        if not self._plugin.IsAcquiring:
            #img = self._plugin.AcquireImageNotShown(id=1)
            #self._plugin.AcquireAndShowImage(mode)
            #img = self._plugin.AcquireImage() # variant
            #img = self._plugin.AcquireFrontImage()  # safe array
            #img = self._plugin.FrontImage  # variant
            #img = self._plugin.AcquireImageShown()
            # img = self._plugin.AcquireDarkSubtractedImage() # variant

            img = self._plugin.AcquireRawImage()  # variant

            if kwargs.get('show', False):
                self._plugin.ShowAcquiredImage()

            return Image(img, name=cameraName, **self._img_params)
        else:
            raise Exception("Camera is busy acquiring...")

    def _set_camera_param(self, name, size, exp_time, binning, camerasize, **kwargs):
        """ Find the TEM camera and set its params. """
        camera_index = self._find_camera(name)
        self._img_params['bit_depth'] = self._plugin.PixelDepth(camera_index)

        self._plugin.CurrentCamera = camera_index

        if self._plugin.IsRetractable:
            if not self._plugin.IsInserted:
                logging.info("Inserting camera %s" % name)
                self._plugin.Insert()
                time.sleep(5)
                if not self._plugin.IsInserted:
                    raise Exception("Could not insert camera!")

        mode = kwargs.get("mode", AcqMode.RECORD)
        self._plugin.SelectCameraParameters(mode)
        self._plugin.Binning = binning
        self._plugin.ExposureTime = exp_time

        speed = kwargs.get("speed", AcqSpeed.SINGLEFRAME)
        self._plugin.Speed = speed

        max_width = camerasize // binning
        max_height = camerasize // binning

        if size == AcqImageSize.FULL:
            self._plugin.CameraLeft = 0
            self._plugin.CameraTop = 0
            self._plugin.CameraRight = max_width
            self._plugin.CameraBottom = max_height
        elif size == AcqImageSize.HALF:
            self._plugin.CameraLeft = int(max_width/4)
            self._plugin.CameraTop = int(max_height/4)
            self._plugin.CameraRight = int(max_width*3/4)
            self._plugin.CameraBottom = int(max_height*3/4)
        elif size == AcqImageSize.QUARTER:
            self._plugin.CameraLeft = int(max_width*3/8)
            self._plugin.CameraTop = int(max_height*3/8)
            self._plugin.CameraRight = int(max_width*3/8 + max_width/4)
            self._plugin.CameraBottom = int(max_height*3/8 + max_height/4)

        self._img_params['width'] = self._plugin.CameraRight - self._plugin.CameraLeft
        self._img_params['height'] = self._plugin.CameraBottom - self._plugin.CameraTop

    def _run_command(self, command, *args):
        #check = 'if(DoesFunctionExist("%s")) Exit(0) else Exit(1)'
        #exists = self._plugin.ExecuteScript(check % command)
        exists = self._plugin.ExecuteScript('DoesFunctionExist("%s")' % command)

        if exists:
            cmd = command % args
            ret = self._plugin.ExecuteScriptFile(cmd)
            if ret:
                raise Exception("Command %s failed" % cmd)


class Image(BaseImage):
    """ Acquired image object. """
    def __init__(self, obj, name=None, **kwargs):
        super().__init__(obj, name, isAdvanced=False, **kwargs)

    @property
    def width(self):
        """ Image width in pixels. """
        return self._kwargs['width']

    @property
    def height(self):
        """ Image height in pixels. """
        return self._kwargs['height']

    @property
    def bit_depth(self):
        """ Bit depth. """
        return self._kwargs['bit_depth']

    @property
    def pixel_type(self):
        """ Image pixels type: uint, int or float. """
        return ImagePixelType.SIGNED_INT.name

    @property
    def data(self):
        """ Returns actual image object as numpy uint16 array. """
        #from comtypes.safearray import safearray_as_ndarray
        #with safearray_as_ndarray:
        #    data = self._img
        #import numpy as np
        data = self._img.astype("uint16")
        data.shape = self.width, self.height

        return data

    def save(self, filename, **kwargs):
        """ Save acquired image to a file.

        :param filename: File path
        :type filename: str
        """
        fmt = os.path.splitext(filename)[1].upper().lstrip(".")
        if fmt == "MRC":
            logging.info("Convert to int16 since MRC does not support int32")
            import mrcfile
            with mrcfile.new(filename) as mrc:
                mrc.set_data(self.data.astype("int16"))
        else:
            raise NotImplementedError("Only mrc format is supported")
