import time
import logging
from datetime import datetime
from ..utils.enums import AcqImageSize, AcqShutterMode, PlateLabelDateFormat, ScreenPosition
from .image import Image


class Acquisition:
    """ Image acquisition functions.

    In order for acquisition to be available TIA (TEM Imaging and Acquisition)
    must be running (even if you are using DigitalMicrograph as the CCD server).

    If it is necessary to update the acquisition object (e.g. when the STEM detector
    selection on the TEM UI has been changed), you have to release and recreate the
    main microscope object. If you do not do so, you keep accessing the same
    acquisition object which will not work properly anymore.
    """
    def __init__(self, client):
        self._client = client
        self._cameras = self._client.get_from_cache("tem.Acquisition.Cameras")
        self._detectors = self._client.get_from_cache("tem.Acquisition.Detectors")
        self._is_advanced = False
        self._prev_shutter_mode = None
        self._eer = False
        self._has_film = None

        if self._client.has_advanced_iface:
            # CSA is supported by Ceta 1, Ceta 2, Falcon 3, Falcon 4
            self._tem_csa = self._client.get_from_cache("tem_adv.Acquisitions.CameraSingleAcquisition")
            if self._client.has("tem_adv.Acquisitions.CameraContinuousAcquisition"):
                # CCA is supported by Ceta 2
                self._tem_cca = self._client.get_from_cache("tem_adv.Acquisitions.CameraContinuousAcquisition")
        else:
                self._tem_cca = None

    @property
    def __has_film(self):
        if self._has_film is None:
            self._has_film = self._client.has("tem.Camera.Stock")
        return self._has_film

    def _find_camera(self, name, recording=False):
        """Find camera object by name. Check adv scripting first. """
        if self._client.has_advanced_iface:
            if recording:
                for cam in self._tem_cca.SupportedCameras:
                    if cam.Name == name:
                        self._is_advanced = True
                        return cam
            for cam in self._tem_csa.SupportedCameras:
                if cam.Name == name:
                    self._is_advanced = True
                    return cam
        for cam in self._cameras:
            if cam.Info.Name == name:
                return cam
        raise KeyError("No camera with name %s. If using standard scripting the "
                       "camera must be selected in the microscope user interface" % name)

    def _find_stem_detector(self, name):
        """Find STEM detector object by name"""
        for stem in self._detectors:
            if stem.Info.Name == name:
                return stem
        raise KeyError("No STEM detector with name %s" % name)

    def _check_binning(self, binning, camera, is_advanced=False, recording=False):
        """ Check if input binning is in SupportedBinnings.

        :param binning: Input binning
        :type binning: int
        :param camera: Camera object
        :param is_advanced: Is this an advanced camera?
        :type is_advanced: bool
        :returns: Binning object
        """
        if is_advanced:
            if recording:
                param = self._tem_cca.CameraSettings.Capabilities
            else:
                param = self._tem_csa.CameraSettings.Capabilities
            for b in param.SupportedBinnings:
                if int(b.Width) == int(binning):
                    return b
        else:
            info = camera.Info
            for b in info.Binnings:
                if int(b) == int(binning):
                    return b

        raise ValueError("Unsupported binning value: %d" % binning)

    def _set_camera_param(self, name, size, exp_time, binning, **kwargs):
        """ Find the TEM camera and set its params. """
        camera = self._find_camera(name, kwargs.get("recording", False))

        if self._is_advanced:
            if not camera.IsInserted:
                camera.Insert()

            if 'recording' in kwargs:
                self._tem_cca.Camera = camera
                settings = self._tem_cca.CameraSettings
                capabilities = settings.Capabilities
                binning = self._check_binning(binning, camera, is_advanced=True, recording=True)
                if hasattr(capabilities, 'SupportsRecording') and capabilities.SupportsRecording:
                    settings.RecordingDuration = kwargs['recording']
                else:
                    raise NotImplementedError("This camera does not support continuous acquisition")

            else:
                self._tem_csa.Camera = camera
                settings = self._tem_csa.CameraSettings
                capabilities = settings.Capabilities
                binning = self._check_binning(binning, camera, is_advanced=True)

            if binning:
                settings.Binning = binning

            settings.ReadoutArea = size

            # Set exposure after binning, since it adjusted
            # automatically when binning is set
            settings.ExposureTime = exp_time

            if 'align_image' in kwargs:
                if capabilities.SupportsDriftCorrection:
                    settings.AlignImage = kwargs['align_image']
                else:
                    raise NotImplementedError("This camera does not support drift correction")

            if 'electron_counting' in kwargs:
                if capabilities.SupportsElectronCounting:
                    settings.ElectronCounting = kwargs['electron_counting']
                else:
                    raise NotImplementedError("This camera does not support electron counting")

            if 'eer' in kwargs and hasattr(capabilities, 'SupportsEER'):
                if capabilities.SupportsEER:
                    self._eer = kwargs['eer']
                    settings.EER = self._eer

                    if self._eer and not settings.ElectronCounting:
                        raise RuntimeError("Electron counting should be enabled when using EER")
                    if self._eer and 'frame_ranges' in kwargs:
                        raise RuntimeError("No frame ranges allowed when using EER")
                else:
                    raise NotImplementedError("This camera does not support EER")

            if 'frame_ranges' in kwargs:  # a list of tuples
                dfd = settings.DoseFractionsDefinition
                dfd.Clear()
                for i in kwargs['frame_ranges']:
                    dfd.AddRange(i[0], i[1])

                now = datetime.now()
                settings.SubPathPattern = name + "_" + now.strftime("%d%m%Y_%H%M%S")
                output = settings.PathToImageStorage + settings.SubPathPattern

                logging.info("Movie of %s frames will be saved to: %s.mrc" % (
                    settings.CalculateNumberOfFrames(), output))
                if not self._eer:
                    logging.info("MRC format can only contain images of up to "
                                 "16-bits per pixel, to get true CameraCounts "
                                 "multiply pixels by PixelToValueCameraCounts "
                                 "factor found in the metadata")

        else:
            info = camera.Info
            settings = camera.AcqParams
            settings.ImageSize = size

            binning = self._check_binning(binning, camera)
            settings.Binning = binning

            if 'correction' in kwargs:
                settings.ImageCorrection = kwargs['correction']
            if 'exposure_mode' in kwargs:
                settings.ExposureMode = kwargs['exposure_mode']
            if 'shutter_mode' in kwargs:
                # Save previous global shutter mode
                self._prev_shutter_mode = (info, info.ShutterMode)
                info.ShutterMode = kwargs['shutter_mode']
            if 'pre_exp_time' in kwargs:
                if kwargs['shutter_mode'] != AcqShutterMode.BOTH:
                    raise RuntimeError("Pre-exposures can only be be done "
                                       "when the shutter mode is set to BOTH")
                settings.PreExposureTime = kwargs['pre_exp_time']
            if 'pre_exp_pause_time' in kwargs:
                if kwargs['shutter_mode'] != AcqShutterMode.BOTH:
                    raise RuntimeError("Pre-exposures can only be be done when "
                                       "the shutter mode is set to BOTH")
                settings.PreExposurePauseTime = kwargs['pre_exp_pause_time']

            # Set exposure after binning, since it adjusted
            # automatically when binning is set
            settings.ExposureTime = exp_time

    def _set_film_param(self, film_text, exp_time):
        """ Set params for plate camera / film. """
        self._cameras.FilmText = film_text.strip()[:96]
        self._cameras.ManualExposureTime = exp_time

    def _acquire(self, cameraName):
        """ Perform actual acquisition.

        :returns: Image object
        """
        acq = self._client.get("tem.Acquisition")
        acq.RemoveAllAcqDevices()
        acq.AddAcqDeviceByName(cameraName)
        imgs = acq.AcquireImages()
        img = imgs[0]

        if self._prev_shutter_mode is not None:
            # restore previous shutter mode
            obj = self._prev_shutter_mode[0]
            old_value = self._prev_shutter_mode[1]
            obj.ShutterMode = old_value

        return Image(img, name=cameraName)

    def _check_prerequisites(self):
        """ Check if buffer cycle or LN filling is
        running before acquisition call. """
        counter = 0
        while counter < 10:
            if self._client.get("tem.Vacuum.PVPRunning"):
                logging.info("Buffer cycle in progress, waiting...\r")
                time.sleep(2)
                counter += 1
            else:
                logging.info("Checking buffer levels...")
                break

        counter = 0
        while counter < 40:
            if (self._client.has("tem.TemperatureControl.TemperatureControlAvailable") and
                    self._client.get("tem.TemperatureControl.DewarsAreBusyFilling")):
                logging.info("Dewars are filling, waiting...\r")
                time.sleep(30)
                counter += 1
            else:
                logging.info("Checking dewars levels...")
                break

    def _acquire_with_tecnaiccd(self, cameraName, size, exp_time,
                                binning, **kwargs):
        if not self._client.has_ccd_iface:
            raise RuntimeError("Tecnai CCD plugin not found, did you "
                               "pass useTecnaiCCD=True to the Microscope() ?")
        else:
            logging.info("Using TecnaiCCD plugin for Gatan camera")
            camerasize = self._find_camera(cameraName).Info.Width  # Get camera size from std scripting
            return self._client._ccd_plugin.acquire_image(
                cameraName, size, exp_time, binning,
                camerasize=camerasize, **kwargs)

    def acquire_tem_image(self, cameraName, size=AcqImageSize.FULL,
                          exp_time=1, binning=1, **kwargs):
        """ Acquire a TEM image.

        :param cameraName: Camera name
        :type cameraName: str
        :param size: Image size (AcqImageSize enum)
        :type size: IntEnum
        :param exp_time: Exposure time in seconds
        :type exp_time: float
        :param binning: Binning factor
        :keyword bool align_image: Whether frame alignment (i.e. drift correction) is to be applied to the final image as well as the intermediate images. Advanced cameras only.
        :keyword bool electron_counting: Use counting mode. Advanced cameras only.
        :keyword bool eer: Use EER mode. Advanced cameras only.
        :keyword list frame_ranges: List of tuple frame ranges that define the intermediate images, e.g. [(1,2), (2,3)]. Advanced cameras only.
        :keyword bool use_tecnaiccd: Use Tecnai CCD plugin to acquire image via Digital Micrograph, only for Gatan cameras. Requires Microscope() initialized with useTecnaiCCD=True
        :returns: Image object

        Usage:
            >>> microscope = Microscope()
            >>> acq = microscope.acquisition
            >>> img = acq.acquire_tem_image("BM-Falcon", AcqImageSize.FULL, exp_time=5.0, binning=1, electron_counting=True, align_image=True)
            >>> img.save("aligned_sum.mrc")
            >>> print(img.width)
            4096
        """
        if kwargs.get("use_tecnaiccd", False):
            return self._acquire_with_tecnaiccd(cameraName, size, exp_time,
                                                binning, **kwargs)

        if kwargs.get("recording", False) and self._tem_cca is None:
            raise NotImplementedError("Recording / continuous acquisition is not available")

        self._set_camera_param(cameraName, size, exp_time, binning, **kwargs)
        if self._is_advanced:
            #while self._tem_csa.IsActive:
            #    time.sleep(0.1)
            self._check_prerequisites()

            if "recording" in kwargs:
                self._tem_cca.Start()
                self._tem_cca.Wait()
                logging.info("Continuous acquisition and offloading job are completed.")
                return None
            else:
                img = self._tem_csa.Acquire()
                self._tem_csa.Wait()
                return Image(img, name=cameraName, isAdvanced=True)

        self._check_prerequisites()
        return self._acquire(cameraName)

    def acquire_stem_image(self, cameraName, size, dwell_time=1E-5,
                           binning=1, **kwargs):
        """ Acquire a STEM image.

        :param cameraName: Camera name
        :type cameraName: str
        :param size: Image size (AcqImageSize enum)
        :type size: IntEnum
        :param dwell_time: Dwell time in seconds. The frame time equals the dwell time times the number of pixels plus some overhead (typically 20%, used for the line flyback)
        :type dwell_time: float
        :param binning: Binning factor
        :type binning: int
        :keyword float brightness: Brightness setting
        :keyword float contrast: Contrast setting
        :returns: Image object
        """
        det = self._find_stem_detector(cameraName)
        info = det.Info

        if 'brightness' in kwargs:
            info.Brightness = kwargs['brightness']
        if 'contrast' in kwargs:
            info.Contrast = kwargs['contrast']

        settings = self._detectors.AcqParams  # self._tem_acq.StemAcqParams
        settings.ImageSize = size

        binning = self._check_binning(binning, det)
        if binning:
            settings.Binning = binning

        settings.DwellTime = dwell_time

        logging.info("Max resolution: %s, %s" % (
            settings.MaxResolution.X,
            settings.MaxResolution.Y))

        self._check_prerequisites()
        return self._acquire(cameraName)

    def acquire_film(self, film_text, exp_time):
        """ Expose a film.

        :param film_text: Film text, 96 symbols
        :type film_text: str
        :param exp_time: Exposure time in seconds
        :type exp_time: float
        """
        if self.__has_film and self._client.get("tem.Camera.Stock") > 0:
            self._cameras.PlateLabelDataType = PlateLabelDateFormat.DDMMYY
            exp_num = self._cameras.ExposureNumber
            self._cameras.ExposureNumber = exp_num + 1
            self._cameras.MainScreen = ScreenPosition.UP
            self._cameras.ScreenDim = True

            self._set_film_param(film_text, exp_time)
            self._cameras.TakeExposure()
            logging.info("Film exposure completed")
        else:
            raise RuntimeError("Plate is not available or stock is empty!")
