import os

from .utils.enums import *
from .base_microscope import BaseImage


class TecnaiCCDCamera:
    """ Main class that uses FEI Tecnai CCD plugin on microscope PC. """
    def __init__(self, microscope):
        self._plugin = microscope._tecnai_ccd
        self._camera_params = dict()

    def _find_camera(self, name):
        """Find camera index by name. """
        for i in range(self._plugin.NumberOfCameras):
            if self._plugin.CameraName == name:
                return i
        raise KeyError("No camera with name %s" % name)

    def acquire_image(self, cameraName, size, exp_time=1, binning=1, **kwargs):
        self._set_camera_param(cameraName, size, exp_time, binning, **kwargs)
        if not self._plugin.IsAcquiring:
            #img = self._plugin.AcquireImageNotShown(id=1)
            #self._plugin.AcquireAndShowImage(mode)
            #img = self._plugin.AcquireImage() # variant
            #img = self._plugin.AcquireFrontImage()  # safe array
            #img = self._plugin.FrontImage  # variant
            #img = self._plugin.AcquireImageShown()
            #self._plugin.ShowAcquiredImage()

            img = self._plugin.AcquireRawImage() # variant
            #img = self._plugin.AcquireDarkSubtractedImage() # variant

            return Image(img, name=cameraName, **self._camera_params)
        else:
            raise Exception("Camera is busy acquiring...")

    def _set_camera_param(self, name, size, exp_time, binning, **kwargs):
        """ Find the TEM camera and set its params. """
        camera_index = self._find_camera(name)
        print("Type: ", self._plugin.Type(camera_index))
        print("Pixel size: ", self._plugin.PixelSize(camera_index))
        self._camera_params['bit_depth'] = self._plugin.PixelDepth(camera_index)

        self._plugin.CurrentCamera = camera_index

        if self._plugin.IsRetractable:
            if not self._plugin.IsInserted:
                self._plugin.Insert()

        mode = kwargs.get("mode", AcqMode.RECORD)
        self._plugin.SelectCameraParameters(mode)
        self._plugin.Binning = binning
        self._plugin.ExposureTime = exp_time

        speed = kwargs.get("speed", AcqSpeed.SINGLEFRAME)
        self._plugin.Speed = speed

        max_width = self._plugin.CameraRight // binning
        max_height = self._plugin.CameraBottom // binning

        if 'left' in kwargs:  # custom size
            try:
                self._check_size(kwargs['left'], kwargs['top'], max_width, max_height)
                self._check_size(kwargs['right'], kwargs['bottom'], max_width, max_height)

                self._plugin.CameraLeft = kwargs['left']
                self._plugin.CameraTop = kwargs['top']
                self._plugin.CameraRight = kwargs['right']
                self._plugin.CameraBottom = kwargs['bottom']
            except KeyError:
                raise Exception("You must specify all params: left, right, top, bottom")
        else:  # account for binning
            self._plugin.CameraLeft = 0
            self._plugin.CameraTop = 0
            self._plugin.CameraRight = max_width // 2
            self._plugin.CameraBottom = max_height // 2

        # Left top is 0,0
        self._camera_params['width'] = self._plugin.CameraRight - self._plugin.CameraLeft
        self._camera_params['height'] = self._plugin.CameraBottom - self._plugin.CameraTop

    def _check_size(self, x, y, max_width, max_height):
        if not (0 <= x <= max_width) or not (0 <= y <= max_height):
            raise Exception("Input image sizes are outside of (0-%d, 0-%d)",
                            max_width, max_height)

    def _run_command(self, command, *args):
        #check = 'if(DoesFunctionExist("%s")) Exit(0) else Exit(1)'
        #exists = self._plugin.ExecuteScript(check % command)
        exists = self._plugin.ExecuteScript('DoesFunctionExist("%s")' % command)

        if exists:
            cmd = command % args
            ret = self._plugin.ExecuteScriptFile(cmd)
            if ret:
                raise Exception("Command %s failed" % cmd)

    def _etc(self):
        self._plugin.OpenShutter(True)  # rw
        self._plugin.LaunchAcquisition(AcqMode.SEARCH)
        self._plugin.StopAcquisition()
        self._plugin.SaveImageInDMFormat(filename="test.dm3")


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
        """ Returns actual image object as numpy int16 array. """
        from comtypes.safearray import safearray_as_ndarray
        with safearray_as_ndarray:
            data = self._img
        return data

    def save(self, filename):
        """ Save acquired image to a file.

        :param filename: File path
        :type filename: str
        """
        fmt = os.path.splitext(filename)[1].upper().replace(".", "")
        if fmt == "MRC":
            print("Convert to int16 since MRC does not support int32")
            import mrcfile
            with mrcfile.new(filename) as mrc:
                mrc.set_data(self.data.astype("int16"))
        else:
            raise NotImplementedError("Only mrc format is supported")
