import os
import logging
import math
import time
from datetime import datetime

from .base_microscope import BaseImage
from .utils.enums import *


ADV_SCR_ERR = "This function is not available in your adv. scripting interface."


class UserDoor:
    """ User door hatch controls. """
    def __init__(self, microscope):
        self._scope = microscope
        self._err_msg = "Door control is unavailable"
        self._tem_door = self._scope.has("tem_adv.UserDoorHatch")

    @property
    def state(self):
        """ Returns door state. """
        if self._tem_door:
            return HatchState(self._scope.get("tem_adv.UserDoorHatch.State")).name
        else:
            raise NotImplementedError(self._err_msg)

    def open(self):
        """ Open the door. """
        if self._tem_door and self._scope.get("tem_adv.UserDoorHatch.IsControlAllowed"):
            self._scope.exec("tem_adv.UserDoorHatch.Open()")
        else:
            raise NotImplementedError(self._err_msg)

    def close(self):
        """ Close the door. """
        if self._tem_door and self._scope.get("tem_adv.UserDoorHatch.IsControlAllowed"):
            self._scope.exec("tem_adv.UserDoorHatch.Close()")
        else:
            raise NotImplementedError(self._err_msg)


class Acquisition:
    """ Image acquisition functions.

    In order for acquisition to be available TIA (TEM Imaging and Acquisition)
    must be running (even if you are using DigitalMicrograph as the CCD server).

    If it is necessary to update the acquisition object (e.g. when the STEM detector
    selection on the TEM UI has been changed), you have to release and recreate the
    main microscope object. If you do not do so, you keep accessing the same
    acquisition object which will not work properly anymore.
    """
    def __init__(self, microscope):
        self._tem = microscope._tem
        self._tem_acq = self._tem.Acquisition
        self._tem_cam = self._tem.Camera
        self._is_advanced = False
        self._has_advanced = microscope._tem_adv is not None
        self._prev_shutter_mode = None
        self._eer = False
        self.__has_film = False

        try:
            _ = self._tem_cam.Stock
            self.__has_film = True
        except:
            pass

        if self._has_advanced:
            self._tem_csa = microscope._tem_adv.Acquisitions.CameraSingleAcquisition

            if hasattr(microscope._tem_adv.Acquisitions, 'CameraContinuousAcquisition'):
                # CCA is supported by Ceta 2
                self._tem_cca = microscope._tem_adv.Acquisitions.CameraContinuousAcquisition
            else:
                self._tem_cca = None

        if getattr(microscope, "_tecnai_ccd_plugin", None):
            self._ccdplugin = microscope._tecnai_ccd_plugin

    def _find_camera(self, name, recording=False):
        """Find camera object by name. Check adv scripting first. """
        if self._has_advanced:
            if recording:
                for cam in self._tem_cca.SupportedCameras:
                    if cam.Name == name:
                        self._is_advanced = True
                        return cam
            for cam in self._tem_csa.SupportedCameras:
                if cam.Name == name:
                    self._is_advanced = True
                    return cam
        for cam in self._tem_acq.Cameras:
            if cam.Info.Name == name:
                return cam
        raise KeyError("No camera with name %s. If using standard scripting the "
                       "camera must be selected in the microscope user interface" % name)

    def _find_stem_detector(self, name):
        """Find STEM detector object by name"""
        for stem in self._tem_acq.Detectors:
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

    def _set_film_param(self, film_text, exp_time, **kwargs):
        """ Set params for plate camera / film. """
        self._tem_cam.FilmText = film_text.strip()[:96]
        self._tem_cam.ManualExposureTime = exp_time

    def _acquire(self, cameraName):
        """ Perform actual acquisition.

        :returns: Image object
        """
        self._tem_acq.RemoveAllAcqDevices()
        self._tem_acq.AddAcqDeviceByName(cameraName)
        imgs = self._tem_acq.AcquireImages()
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
        tc = self._tem.TemperatureControl
        counter = 0
        while counter < 10:
            if self._tem.Vacuum.PVPRunning:
                logging.info("Buffer cycle in progress, waiting...\r")
                time.sleep(2)
                counter += 1
            else:
                logging.info("Checking buffer levels...")
                break

        counter = 0
        while counter < 40:
            if tc.TemperatureControlAvailable and tc.DewarsAreBusyFilling:
                logging.info("Dewars are filling, waiting...\r")
                time.sleep(30)
                counter += 1
            else:
                logging.info("Checking dewars levels...")
                break

    def _acquire_with_tecnaiccd(self, cameraName, size, exp_time,
                                binning, **kwargs):
        if not hasattr(self, "_ccdplugin"):
            raise RuntimeError("Tecnai CCD plugin not found, did you "
                               "pass useTecnaiCCD=True to Microscope() ?")
        else:
            logging.info("Using TecnaiCCD plugin for Gatan camera")
            camerasize = self._find_camera(cameraName).Info.Width  # Get camera size from std scripting
            return self._ccdplugin.acquire_image(cameraName, size, exp_time, binning,
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

        settings = self._tem_acq.Detectors.AcqParams  # self._tem_acq.StemAcqParams
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

    def acquire_film(self, film_text, exp_time, **kwargs):
        """ Expose a film.

        :param film_text: Film text, 96 symbols
        :type film_text: str
        :param exp_time: Exposure time in seconds
        :type exp_time: float
        """
        if self.__has_film and self._tem_cam.Stock > 0:
            self._tem_cam.PlateLabelDataType = PlateLabelDateFormat.DDMMYY
            exp_num = self._tem_cam.ExposureNumber
            self._tem_cam.ExposureNumber = exp_num + 1
            self._tem_cam.MainScreen = ScreenPosition.UP
            self._tem_cam.ScreenDim = True

            self._set_film_param(film_text, exp_time, **kwargs)
            self._tem_cam.TakeExposure()
            logging.info("Film exposure completed")
        else:
            raise RuntimeError("Plate is not available or stock is empty!")


class Detectors:
    """ CCD/DDD, film/plate and STEM detectors. """
    def __init__(self, microscope):
        self._tem_acq = microscope._tem.Acquisition
        self._tem_cam = microscope._tem.Camera
        self._has_advanced = microscope._tem_adv is not None
        self.__has_film = False

        try:
            _ = self._tem_cam.Stock
            self.__has_film = True
        except:
            logging.info("No film/plate device detected.")

        if self._has_advanced:
            # CSA is supported by Ceta 1, Ceta 2, Falcon 3, Falcon 4
            self._tem_csa = microscope._tem_adv.Acquisitions.CameraSingleAcquisition
            if hasattr(microscope._tem_adv.Acquisitions, 'CameraContinuousAcquisition'):
                # CCA is supported by Ceta 2
                self._tem_cca = microscope._tem_adv.Acquisitions.CameraContinuousAcquisition
                self._cca_cameras = [c.Name for c in self._tem_cca.SupportedCameras]
            else:
                self._tem_cca = None
                logging.info("Continuous acquisition not supported.")

    @property
    def cameras(self):
        """ Returns a dict with parameters for all cameras. """
        self._cameras = dict()
        for cam in self._tem_acq.Cameras:
            info = cam.Info
            param = cam.AcqParams
            name = info.Name
            self._cameras[name] = {
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
        if not self._has_advanced:
            return self._cameras

        for cam in self._tem_csa.SupportedCameras:
            self._tem_csa.Camera = cam
            param = self._tem_csa.CameraSettings.Capabilities
            self._cameras[cam.Name] = {
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
                self._cameras[cam.Name].update({
                    "supports_recording": getattr(param, 'SupportsRecording', False)
                })

        return self._cameras

    @property
    def stem_detectors(self):
        """ Returns a dict with STEM detectors parameters. """
        self._stem_detectors = dict()
        for det in self._tem_acq.Detectors:
            info = det.Info
            name = info.Name
            self._stem_detectors[name] = {
                "type": "STEM_DETECTOR",
                "binnings": [int(b) for b in info.Binnings]
            }
        return self._stem_detectors

    @property
    def screen(self):
        """ Fluorescent screen position. (read/write)"""
        return ScreenPosition(self._tem_cam.MainScreen).name

    @screen.setter
    def screen(self, value):
        self._tem_cam.MainScreen = value

    @property
    def film_settings(self):
        """ Returns a dict with film settings.
        Note: The plate camera has become obsolete with Win7 so
        most of the existing functions are no longer supported.
        """
        if self.__has_film:
            return {
                "stock": self._tem_cam.Stock,  # Int
                "exposure_time": self._tem_cam.ManualExposureTime,
                "film_text": self._tem_cam.FilmText,
                "exposure_number": self._tem_cam.ExposureNumber,
                "user_code": self._tem_cam.Usercode,  # 3 digits
                "screen_current": self._tem_cam.ScreenCurrent * 1e9  # check if works without film
            }
        else:
            logging.info("No film/plate device detected.")


class Temperature:
    """ LN dewars and temperature controls. """
    def __init__(self, microscope):
        self._scope = microscope
        self._err_msg = "TemperatureControl is not available"
        self._tem_tmpctrl = self._scope.has("tem.TemperatureControl")
        self._tem_tmpctrl_adv = self._scope.has("tem_adv.TemperatureControl")

    @property
    def is_available(self):
        """ Status of the temperature control. Should be always False on Tecnai instruments. """
        if self._tem_tmpctrl:
            return self._scope.get("tem.TemperatureControl.TemperatureControlAvailable")
        else:
            raise RuntimeError(self._err_msg)

    def force_refill(self):
        """ Forces LN refill if the level is below 70%, otherwise returns an error.
        Note: this function takes considerable time to execute.
        """
        if self.is_available:
            self._scope.exec("tem.TemperatureControl.ForceRefill()")
        elif self._tem_tmpctrl_adv:
            return self._scope.exec("tem_adv.TemperatureControl.RefillAllDewars()")
        else:
            raise RuntimeError(self._err_msg)

    def dewar_level(self, dewar):
        """ Returns the LN level (%) in a dewar.

        :param dewar: Dewar name (RefrigerantDewar enum)
        :type dewar: IntEnum
        """
        if self.is_available:
            return self._scope.exec("tem.TemperatureControl.RefrigerantLevel()", dewar)
        else:
            raise RuntimeError(self._err_msg)

    @property
    def is_dewar_filling(self):
        """ Returns TRUE if any of the dewars is currently busy filling. """
        if self.is_available:
            return self._scope.get("tem.TemperatureControl.DewarsAreBusyFilling")
        elif self._tem_tmpctrl_adv:
            return self._scope.get("tem_adv.TemperatureControl.IsAnyDewarFilling")
        else:
            raise RuntimeError(self._err_msg)

    @property
    def dewars_time(self):
        """ Returns remaining time (seconds) until the next dewar refill.
        Returns -1 if no refill is scheduled (e.g. All room temperature, or no
        dewar present).
        """
        # TODO: check if returns -60 at room temperature
        if self.is_available:
            return self._scope.get("tem.TemperatureControl.DewarsRemainingTime")
        else:
            raise RuntimeError(self._err_msg)

    @property
    def temp_docker(self):
        """ Returns Docker temperature in Kelvins. """
        if self._tem_tmpctrl_adv:
            return self._scope.get("tem_adv.TemperatureControl.AutoloaderCompartment.DockerTemperature")
        else:
            raise NotImplementedError(ADV_SCR_ERR)

    @property
    def temp_cassette(self):
        """ Returns Cassette gripper temperature in Kelvins. """
        if self._tem_tmpctrl_adv:
            return self._scope.get("tem_adv.TemperatureControl.AutoloaderCompartment.CassetteTemperature")
        else:
            raise NotImplementedError(ADV_SCR_ERR)

    @property
    def temp_cartridge(self):
        """ Returns Cartridge gripper temperature in Kelvins. """
        if self._tem_tmpctrl_adv:
            return self._scope.get("tem_adv.TemperatureControl.AutoloaderCompartment.CartridgeTemperature")
        else:
            raise NotImplementedError(ADV_SCR_ERR)

    @property
    def temp_holder(self):
        """ Returns Holder temperature in Kelvins. """
        if self._tem_tmpctrl_adv:
            return self._scope.get("tem_adv.TemperatureControl.ColumnCompartment.HolderTemperature")
        else:
            raise NotImplementedError(ADV_SCR_ERR)


class Autoloader:
    """ Sample loading functions. """
    def __init__(self, microscope):
        self._scope = microscope
        self._err_msg = "Autoloader is not available"
        self._tem_autoloader_adv = self._scope.has("tem_adv.AutoLoader")

    @property
    def is_available(self):
        """ Status of the autoloader. Should be always False on Tecnai instruments. """
        return self._scope.get("tem.AutoLoader.AutoLoaderAvailable")

    @property
    def number_of_slots(self):
        """ The number of slots in a cassette. """
        if self.is_available:
            return self._scope.get("tem.AutoLoader.NumberOfCassetteSlots")
        else:
            raise RuntimeError(self._err_msg)

    def load_cartridge(self, slot):
        """ Loads the cartridge in the given slot into the microscope.

        :param slot: Slot number
        :type slot: int
        """
        if self.is_available:
            total = self.number_of_slots
            slot = int(slot)
            if slot > total:
                raise ValueError("Only %s slots are available" % total)
            if self.slot_status(slot) != CassetteSlotStatus.OCCUPIED.name:
                raise RuntimeError("Slot %d is not occupied" % slot)
            self._scope.exec("tem.AutoLoader.LoadCartridge()", slot)
        else:
            raise RuntimeError(self._err_msg)

    def unload_cartridge(self):
        """ Unloads the cartridge currently in the microscope and puts it back into its
        slot in the cassette.
        """
        if self.is_available:
            self._scope.exec("tem.AutoLoader.UnloadCartridge()")
        else:
            raise RuntimeError(self._err_msg)

    def run_inventory(self):
        """ Performs an inventory of the cassette.
        Note: This function takes considerable time to execute.
        """
        # TODO: check if cassette is present
        if self.is_available:
            self._scope.exec("tem.AutoLoader.PerformCassetteInventory()")
        else:
            raise RuntimeError(self._err_msg)

    def slot_status(self, slot):
        """ The status of the slot specified.

        :param slot: Slot number
        :type slot: int
        """
        if self.is_available:
            total = self.number_of_slots
            if slot > total:
                raise ValueError("Only %s slots are available" % total)
            status = self._scope.exec("tem.AutoLoader.SlotStatus()", int(slot))
            return CassetteSlotStatus(status).name
        else:
            raise RuntimeError(self._err_msg)

    def undock_cassette(self):
        """ Moves the cassette from the docker to the capsule. """
        if self._tem_autoloader_adv:
            if self.is_available:
                self._scope.exec("tem_adv.AutoLoader.UndockCassette()")
            else:
                raise RuntimeError(self._err_msg)
        else:
            raise NotImplementedError(ADV_SCR_ERR)

    def dock_cassette(self):
        """ Moves the cassette from the capsule to the docker. """
        if self._tem_autoloader_adv:
            if self.is_available:
                self._scope.exec("tem_adv.AutoLoader.DockCassette()")
            else:
                raise RuntimeError(self._err_msg)
        else:
            raise NotImplementedError(ADV_SCR_ERR)

    def initialize(self):
        """ Initializes / Recovers the Autoloader for further use. """
        if self._tem_autoloader_adv:
            if self.is_available:
                self._scope.exec("tem_adv.AutoLoader.Initialize()")
            else:
                raise RuntimeError(self._err_msg)
        else:
            raise NotImplementedError(ADV_SCR_ERR)

    def buffer_cycle(self):
        """ Synchronously runs the Autoloader buffer cycle. """
        if self._tem_autoloader_adv:
            if self.is_available:
                self._scope.exec("tem_adv.AutoLoader.BufferCycle()")
            else:
                raise RuntimeError(self._err_msg)
        else:
            raise NotImplementedError(ADV_SCR_ERR)


class Stage:
    """ Stage functions. """
    def __init__(self, microscope):
        self._scope = microscope
        self._err_msg = "Stage is not ready"

    def _from_dict(self, **values):
        axes = 0
        position = self._scope.get("tem.Stage.Position")
        for key, value in values.items():
            if key not in 'xyzab':
                raise ValueError("Unexpected axis: %s" % key)
            attr_name = key.upper()
            setattr(position, attr_name, float(value))
            axes |= getattr(StageAxes, attr_name)
        return position, axes

    def _beta_available(self):
        return self.limits['b']['unit'] != MeasurementUnitType.UNKNOWN.name

    def _change_position(self, direct=False, tries=5, **kwargs):
        attempt = 0
        while attempt < tries:
            if self._scope.get("tem.Stage.Status") != StageStatus.READY:
                logging.info("Stage is not ready, retrying...")
                tries += 1
                time.sleep(1)
            else:
                # convert units to meters and radians
                new_pos = dict()
                for axis in 'xyz':
                    if axis in kwargs:
                        new_pos.update({axis: kwargs[axis] * 1e-6})
                for axis in 'ab':
                    if axis in kwargs:
                        new_pos.update({axis: math.radians(kwargs[axis])})

                speed = kwargs.get("speed", None)
                if speed is not None and not (0.0 <= speed <= 1.0):
                    raise ValueError("Speed must be within 0.0-1.0 range")

                if 'b' in new_pos and not self._beta_available():
                    raise KeyError("B-axis is not available")

                limits = self.limits
                for key, value in new_pos.items():
                    if value < limits[key]['min'] or value > limits[key]['max']:
                        raise ValueError('Stage position %s=%s is out of range' % (value, key))

                # X and Y - 1000 to + 1000(micrometers)
                # Z - 375 to 375(micrometers)
                # a - 80 to + 80(degrees)
                # b - 29.7 to + 29.7(degrees)

                new_pos, axes = self._from_dict(**new_pos)
                if not direct:
                    self._scope.exec("tem.Stage.MoveTo()", new_pos, axes)
                else:
                    if speed is not None:
                        self._scope.exec("tem.Stage.GoToWithSpeed()", new_pos, axes, speed)
                    else:
                        self._scope.exec("tem.Stage.GoTo()", new_pos, axes)

                break
        else:
            raise RuntimeError(self._err_msg)

    @property
    def status(self):
        """ The current state of the stage. """
        return StageStatus(self._scope.get("tem.Stage.Status")).name

    @property
    def holder(self):
        """ The current specimen holder type. """
        return StageHolderType(self._scope.get("tem.Stage.Holder")).name

    @property
    def position(self):
        """ The current position of the stage (x,y,z in um and a,b in degrees). """
        pos = self._scope.get("tem.Stage.Position")
        result = {key.lower(): getattr(pos, key) * 1e6 for key in 'XYZ'}

        keys = 'AB' if self._beta_available() else 'A'
        result.update({key.lower(): math.degrees(getattr(pos, key)) for key in keys})

        return result

    def go_to(self, **kwargs):
        """ Makes the holder directly go to the new position by moving all axes
        simultaneously. Keyword args can be x,y,z,a or b.

        :keyword float speed: fraction of the standard speed setting (max 1.0)
        """
        self._change_position(direct=True, **kwargs)

    def move_to(self, **kwargs):
        """ Makes the holder safely move to the new position.
        Keyword args can be x,y,z,a or b.
        """
        kwargs['speed'] = None
        self._change_position(**kwargs)

    @property
    def limits(self):
        """ Returns a dict with stage move limits. """
        result = dict()
        for axis in 'xyzab':
            data = self._scope.exec("tem.Stage.AxisData()", StageAxes[axis.upper()])
            result[axis] = {
                'min': data.MinPos,
                'max': data.MaxPos,
                'unit': MeasurementUnitType(data.UnitType).name
            }
        return result


class PiezoStage:
    """ Piezo stage functions. """
    def __init__(self, microscope):
        self._scope = microscope
        try:
            self._tem_pstage = self._scope.has("tem_adv.PiezoStage")
            _ = self._scope.has("tem_adv.PiezoStage.HighResolution")
        except:
            logging.info("PiezoStage interface is not available.")

    @property
    def position(self):
        """ The current position of the piezo stage (x,y,z in um). """
        pos = self._scope.get("tem_adv.PiezoStage.CurrentPosition")
        return {key: getattr(pos, key.upper()) * 1e6 for key in 'xyz'}

    @property
    def position_range(self):
        """ Return min and max positions. """
        return self._scope.exec("tem_adv.PiezoStage.GetPositionRange()")

    @property
    def velocity(self):
        """ Returns a dict with stage velocities. """
        pos = self._scope.get("tem_adv.PiezoStage.CurrentJogVelocity")
        return {key: getattr(pos, key.upper()) for key in 'xyz'}


class Vacuum:
    """ Vacuum functions. """
    def __init__(self, microscope):
        self._scope = microscope

    @property
    def status(self):
        """ Status of the vacuum system. """
        return VacuumStatus(self._scope.get("tem.Vacuum.Status")).name

    @property
    def is_buffer_running(self):
        """ Checks whether the prevacuum pump is currently running
        (consequences: vibrations, exposure function blocked
        or should not be called).
        """
        return self._scope.get("tem.Vacuum.PVPRunning")

    @property
    def is_column_open(self):
        """ The status of the column valves. """
        return self._scope.get("tem.Vacuum.ColumnValvesOpen")

    @property
    def gauges(self):
        """ Returns a dict with vacuum gauges information.
        Pressure values are in Pascals.
        """
        gauges = {}
        for g in self._scope.get("tem.Vacuum.Gauges"):
            # g.Read()
            if g.Status == GaugeStatus.UNDEFINED:
                # set manually if undefined, otherwise fails
                pressure_level = GaugePressureLevel.UNDEFINED
            else:
                pressure_level = GaugePressureLevel(g.PressureLevel).name

            gauges[g.Name] = {
                "status": GaugeStatus(g.Status).name,
                "pressure": g.Pressure,
                "trip_level": pressure_level
            }
        return gauges

    def column_open(self):
        """ Open column valves. """
        self._scope.set("tem.Vacuum.ColumnValvesOpen", True)

    def column_close(self):
        """ Close column valves. """
        self._scope.set("tem.Vacuum.ColumnValvesOpen", False)

    def run_buffer_cycle(self):
        """ Runs a pumping cycle to empty the buffer. """
        self._scope.exec("tem.Vacuum.RunBufferCycle()")


class Optics:
    """ Projection, Illumination functions. """
    def __init__(self, microscope):
        self._scope = microscope
        self.illumination = Illumination(microscope)
        self.projection = Projection(microscope)

    @property
    def screen_current(self):
        """ The current measured on the fluorescent screen (units: nanoAmperes). """
        return self._scope.get("tem.Camera.ScreenCurrent") * 1e9

    @property
    def is_beam_blanked(self):
        """ Status of the beam blanker. """
        return self._scope.get("tem.Illumination.BeamBlanked")

    @property
    def is_shutter_override_on(self):
        """ Determines the state of the shutter override function.
        WARNING: Do not leave the Shutter override on when stopping the script.
        The microscope operator will be unable to have a beam come down and has
        no separate way of seeing that it is blocked by the closed microscope shutter.
        """
        return self._scope.get("tem.BlankerShutter.ShutterOverrideOn")

    @property
    def is_autonormalize_on(self):
        """ Status of the automatic normalization procedures performed by
        the TEM microscope. Normally they are active, but for scripting it can be
        convenient to disable them temporarily.
        """
        return self._scope.get("tem.AutoNormalizeEnabled")

    def beam_blank(self):
        """ Activates the beam blanker. """
        self._scope.set("tem.Illumination.BeamBlanked", True)
        logging.warning("Falcon protector might delay blanker response")

    def beam_unblank(self):
        """ Deactivates the beam blanker. """
        self._scope.set("tem.Illumination.BeamBlanked", False)
        logging.warning("Falcon protector might delay blanker response")

    def normalize_all(self):
        """ Normalize all lenses. """
        self._scope.exec("tem.NormalizeAll()")

    def normalize(self, mode):
        """ Normalize condenser or projection lens system.
        :param mode: Normalization mode (ProjectionNormalization or IlluminationNormalization enum)
        :type mode: IntEnum
        """
        if mode in ProjectionNormalization:
            self._scope.exec("tem.Projection.Normalize()", mode)
        elif mode in IlluminationNormalization:
            self._scope.exec("tem.Illumination.Normalize()", mode)
        else:
            raise ValueError("Unknown normalization mode: %s" % mode)


class Stem:
    """ STEM functions. """
    def __init__(self, microscope):
        self._scope = microscope
        self._err_msg = "Microscope not in STEM mode"

    @property
    def is_available(self):
        """ Returns whether the microscope has a STEM system or not. """
        return self._scope.get("tem.InstrumentModeControl.StemAvailable")

    def enable(self):
        """ Switch to STEM mode."""
        if self.is_available:
            self._scope.set("tem.InstrumentModeControl.InstrumentMode", InstrumentMode.STEM)
        else:
            raise RuntimeError(self._err_msg)

    def disable(self):
        """ Switch back to TEM mode. """
        self._scope.set("tem.InstrumentModeControl.InstrumentMode", InstrumentMode.TEM)

    @property
    def magnification(self):
        """ The magnification value in STEM mode. (read/write)"""
        if self._scope.get("tem.InstrumentModeControl.InstrumentMode") == InstrumentMode.STEM:
            return self._scope.get("tem.Illumination.StemMagnification")
        else:
            raise RuntimeError(self._err_msg)

    @magnification.setter
    def magnification(self, mag):
        if self._scope.get("tem.InstrumentModeControl.InstrumentMode") == InstrumentMode.STEM:
            self._scope.set("tem.Illumination.StemMagnification", float(mag))
        else:
            raise RuntimeError(self._err_msg)

    @property
    def rotation(self):
        """ The STEM rotation angle (in mrad). (read/write)"""
        if self._scope.get("tem.InstrumentModeControl.InstrumentMode") == InstrumentMode.STEM:
            return self._scope.get("tem.Illumination.StemRotation") * 1e3
        else:
            raise RuntimeError(self._err_msg)

    @rotation.setter
    def rotation(self, rot):
        if self._scope.get("tem.InstrumentModeControl.InstrumentMode") == InstrumentMode.STEM:
            self._scope.set("tem.Illumination.StemRotation", float(rot) * 1e-3)
        else:
            raise RuntimeError(self._err_msg)

    @property
    def scan_field_of_view(self):
        """ STEM full scan field of view. (read/write)"""
        if self._scope.get("tem.InstrumentModeControl.InstrumentMode") == InstrumentMode.STEM:
            return (self._scope.get("tem.Illumination.StemFullScanFieldOfView.X"),
                    self._scope.get("tem.Illumination.StemFullScanFieldOfView.Y"))
        else:
            raise RuntimeError(self._err_msg)

    @scan_field_of_view.setter
    def scan_field_of_view(self, values):
        if self._scope.get("tem.InstrumentModeControl.InstrumentMode") == InstrumentMode.STEM:
            self._scope.set("tem.Illumination.StemFullScanFieldOfView", values, vector=True)
        else:
            raise RuntimeError(self._err_msg)


class Illumination:
    """ Illumination functions. """
    def __init__(self, microscope):
        self._scope = microscope

    @property
    def spotsize(self):
        """ Spotsize number, usually 1 to 11. (read/write)"""
        return self._scope.get("tem.Illumination.SpotsizeIndex")

    @spotsize.setter
    def spotsize(self, value):
        if not (1 <= int(value) <= 11):
            raise ValueError("%s is outside of range 1-11" % value)
        self._scope.set("tem.Illumination.SpotsizeIndex", int(value))

    @property
    def intensity(self):
        """ Intensity / C2 condenser lens value. (read/write)"""
        return self._scope.get("tem.Illumination.Intensity")

    @intensity.setter
    def intensity(self, value):
        if not (0.0 <= value <= 1.0):
            raise ValueError("%s is outside of range 0.0-1.0" % value)
        self._scope.set("tem.Illumination.Intensity", float(value))

    @property
    def intensity_zoom(self):
        """ Intensity zoom. Set to False to disable. (read/write)"""
        return self._scope.get("tem.Illumination.IntensityZoomEnabled")

    @intensity_zoom.setter
    def intensity_zoom(self, value):
        self._scope.set("tem.Illumination.IntensityZoomEnabled", bool(value))

    @property
    def intensity_limit(self):
        """ Intensity limit. Set to False to disable. (read/write)"""
        return self._scope.get("tem.Illumination.IntensityLimitEnabled")

    @intensity_limit.setter
    def intensity_limit(self, value):
        self._scope.set("tem.Illumination.IntensityLimitEnabled", bool(value))

    @property
    def beam_shift(self):
        """ Beam shift X and Y in um. (read/write)"""
        return (self._scope.get("tem.Illumination.Shift.X") * 1e6,
                self._scope.get("tem.Illumination.Shift.Y") * 1e6)

    @beam_shift.setter
    def beam_shift(self, value):
        new_value = (value[0] * 1e-6, value[1] * 1e-6)
        self._scope.set("tem.Illumination.Shift", new_value, vector=True)

    @property
    def rotation_center(self):
        """ Rotation center X and Y in mrad. (read/write)
            Depending on the scripting version,
            the values might need scaling by 6.0 to get mrads.
        """
        return (self._scope.get("tem.Illumination.RotationCenter.X") * 1e3,
                self._scope.get("tem.Illumination.RotationCenter.Y") * 1e3)

    @rotation_center.setter
    def rotation_center(self, value):
        new_value = (value[0] * 1e-3, value[1] * 1e-3)
        self._scope.set("tem.Illumination.RotationCenter", new_value, vector=True)

    @property
    def condenser_stigmator(self):
        """ C2 condenser stigmator X and Y. (read/write)"""
        return (self._scope.get("tem.Illumination.CondenserStigmator.X"),
                self._scope.get("tem.Illumination.CondenserStigmator.Y"))

    @condenser_stigmator.setter
    def condenser_stigmator(self, value):
        self._scope.set("tem.Illumination.CondenserStigmator", value, vector=True, limits=(-1.0, 1.0))

    @property
    def illuminated_area(self):
        """ Illuminated area. Works only on 3-condenser lens systems. (read/write)"""
        if self._scope.get("tem.Configuration.CondenserLensSystem") == CondenserLensSystem.THREE_CONDENSER_LENSES:
            return self._scope.get("tem.Illumination.IlluminatedArea")
        else:
            raise NotImplementedError("Illuminated area exists only on 3-condenser lens systems.")

    @illuminated_area.setter
    def illuminated_area(self, value):
        if self._scope.get("tem.Configuration.CondenserLensSystem") == CondenserLensSystem.THREE_CONDENSER_LENSES:
            self._scope.set("tem.Illumination.IlluminatedArea", float(value))
        else:
            raise NotImplementedError("Illuminated area exists only on 3-condenser lens systems.")

    @property
    def probe_defocus(self):
        """ Probe defocus. Works only on 3-condenser lens systems. (read/write)"""
        if self._scope.get("tem.Configuration.CondenserLensSystem") == CondenserLensSystem.THREE_CONDENSER_LENSES:
            return self._scope.get("tem.Illumination.ProbeDefocus")
        else:
            raise NotImplementedError("Probe defocus exists only on 3-condenser lens systems.")

    @probe_defocus.setter
    def probe_defocus(self, value):
        if self._scope.get("tem.Configuration.CondenserLensSystem") == CondenserLensSystem.THREE_CONDENSER_LENSES:
            self._scope.set("tem.Illumination.ProbeDefocus", float(value))
        else:
            raise NotImplementedError("Probe defocus exists only on 3-condenser lens systems.")

    #TODO: check if the illum. mode is probe?
    @property
    def convergence_angle(self):
        """ Convergence angle. Works only on 3-condenser lens systems. (read/write)"""
        if self._scope.get("tem.Configuration.CondenserLensSystem") == CondenserLensSystem.THREE_CONDENSER_LENSES:
            return self._scope.get("tem.Illumination.ConvergenceAngle")
        else:
            raise NotImplementedError("Convergence angle exists only on 3-condenser lens systems.")

    @convergence_angle.setter
    def convergence_angle(self, value):
        if self._scope.get("tem.Configuration.CondenserLensSystem") == CondenserLensSystem.THREE_CONDENSER_LENSES:
            self._scope.set("tem.Illumination.ConvergenceAngle", float(value))
        else:
            raise NotImplementedError("Convergence angle exists only on 3-condenser lens systems.")

    @property
    def C3ImageDistanceParallelOffset(self):
        """ C3 image distance parallel offset. Works only on 3-condenser lens systems. (read/write)"""
        if self._scope.get("tem.Configuration.CondenserLensSystem") == CondenserLensSystem.THREE_CONDENSER_LENSES:
            return self._scope.get("tem.Illumination.C3ImageDistanceParallelOffset")
        else:
            raise NotImplementedError("C3ImageDistanceParallelOffset exists only on 3-condenser lens systems.")

    @C3ImageDistanceParallelOffset.setter
    def C3ImageDistanceParallelOffset(self, value):
        if self._scope.get("tem.Configuration.CondenserLensSystem") == CondenserLensSystem.THREE_CONDENSER_LENSES:
            self._scope.set("tem.Illumination.C3ImageDistanceParallelOffset", float(value))
        else:
            raise NotImplementedError("C3ImageDistanceParallelOffset exists only on 3-condenser lens systems.")

    @property
    def mode(self):
        """ Illumination mode: microprobe or nanoprobe. (read/write)"""
        return IlluminationMode(self._scope.get("tem.Illumination.Mode")).name

    @mode.setter
    def mode(self, value):
        self._scope.set("tem.Illumination.Mode", value)

    @property
    def dark_field(self):
        """ Dark field mode: cartesian, conical or off. (read/write)"""
        return DarkFieldMode(self._scope.get("tem.Illumination.DFMode")).name

    @dark_field.setter
    def dark_field(self, value):
        self._scope.set("tem.Illumination.DFMode", value)

    @property
    def condenser_mode(self):
        """ Mode of the illumination system: parallel or probe. (read/write)"""
        if self._scope.get("tem.Configuration.CondenserLensSystem") == CondenserLensSystem.THREE_CONDENSER_LENSES:
            return CondenserMode(self._scope.get("tem.Illumination.CondenserMode")).name
        else:
            raise NotImplementedError("Condenser mode exists only on 3-condenser lens systems.")

    @condenser_mode.setter
    def condenser_mode(self, value):
        if self._scope.get("tem.Configuration.CondenserLensSystem") == CondenserLensSystem.THREE_CONDENSER_LENSES:
            self._scope.set("tem.Illumination.CondenserMode", value)
        else:
            raise NotImplementedError("Condenser mode can be changed only on 3-condenser lens systems.")

    @property
    def beam_tilt(self):
        """ Dark field beam tilt relative to the origin stored at
        alignment time. Only operational if dark field mode is active.
        Units: mrad, either in Cartesian (x,y) or polar (conical)
        tilt angles. The accuracy of the beam tilt physical units
        depends on a calibration of the tilt angles. (read/write)
        """
        mode = self._scope.get("tem.Illumination.DFMode")
        tilt = self._scope.get("tem.Illumination.Tilt")
        if mode == DarkFieldMode.CONICAL:
            return tilt[0] * 1e3 * math.cos(tilt[1]), tilt[0] * 1e3 * math.sin(tilt[1])
        elif mode == DarkFieldMode.CARTESIAN:
            return tilt * 1e3
        else:
            return 0.0, 0.0  # Microscope might return nonsense if DFMode is OFF

    @beam_tilt.setter
    def beam_tilt(self, tilt):
        mode = self._scope.get("tem.Illumination.DFMode")
        tilt[0] *= 1e-3
        tilt[1] *= 1e-3
        if tilt[0] == 0.0 and tilt[1] == 0.0:
            self._scope.set("tem.Illumination.Tilt", (0.0, 0.0))
            self._scope.set("tem.Illumination.DFMode", DarkFieldMode.OFF)
        elif mode == DarkFieldMode.CONICAL:
            self._scope.set("tem.Illumination.Tilt", (math.sqrt(tilt[0] ** 2 + tilt[1] ** 2), math.atan2(tilt[1], tilt[0])))
        elif mode == DarkFieldMode.OFF:
            self._scope.set("tem.Illumination.DFMode", DarkFieldMode.CARTESIAN)
            self._scope.set("tem.Illumination.Tilt", tilt)
        else:
            self._scope.set("tem.Illumination.Tilt", tilt)


class Projection:
    """ Projection system functions. """
    def __init__(self, microscope):
        self._scope = microscope
        self._err_msg = "Microscope is not in diffraction mode"
        #self.magnification_index = self._tem_projection.MagnificationIndex
        #self.camera_length_index = self._tem_projection.CameraLengthIndex

    @property
    def focus(self):
        """ Absolute focus value. (read/write)"""
        return self._scope.get("tem.Projection.Focus")

    @focus.setter
    def focus(self, value):
        if not (-1.0 <= value <= 1.0):
            raise ValueError("%s is outside of range -1.0 to 1.0" % value)

        self._scope.set("tem.Projection.Focus", float(value))

    @property
    def magnification(self):
        """ The reference magnification value (screen up setting)."""
        if self._scope.get("tem.Projection.Mode") == ProjectionMode.IMAGING:
            return self._scope.get("tem.Projection.Magnification")
        else:
            raise RuntimeError(self._err_msg)

    @property
    def camera_length(self):
        """ The reference camera length in m (screen up setting). """
        if self._scope.get("tem.Projection.Mode") == ProjectionMode.DIFFRACTION:
            return self._scope.get("tem.Projection.CameraLength")
        else:
            raise RuntimeError(self._err_msg)

    @property
    def image_shift(self):
        """ Image shift in um. (read/write)"""
        return (self._scope.get("tem.Projection.ImageShift.X") * 1e6,
                self._scope.get("tem.Projection.ImageShift.Y") * 1e6)

    @image_shift.setter
    def image_shift(self, value):
        new_value = (value[0] * 1e-6, value[1] * 1e-6)
        self._scope.set("tem.Projection.ImageShift", new_value, vector=True)

    @property
    def image_beam_shift(self):
        """ Image shift with beam shift compensation in um. (read/write)"""
        return (self._scope.get("tem.Projection.ImageBeamShift.X") * 1e6,
                self._scope.get("tem.Projection.ImageBeamShift.Y") * 1e6)

    @image_beam_shift.setter
    def image_beam_shift(self, value):
        new_value = (value[0] * 1e-6, value[1] * 1e-6)
        self._scope.set("tem.Projection.ImageBeamShift", new_value, vector=True)

    @property
    def image_beam_tilt(self):
        """ Beam tilt with diffraction shift compensation in mrad. (read/write)"""
        return (self._scope.get("tem.Projection.ImageBeamTilt.X") * 1e3,
                self._scope.get("tem.Projection.ImageBeamTilt.Y") * 1e3)

    @image_beam_tilt.setter
    def image_beam_tilt(self, value):
        new_value = (value[0] * 1e-3, value[1] * 1e-3)
        self._scope.set("tem.Projection.ImageBeamTilt", new_value, vector=True)

    @property
    def diffraction_shift(self):
        """ Diffraction shift in mrad. (read/write)"""
        #TODO: 180/pi*value = approx number in TUI
        return (self._scope.get("tem.Projection.DiffractionShift.X") * 1e3,
                self._scope.get("tem.Projection.DiffractionShift.Y") * 1e3)

    @diffraction_shift.setter
    def diffraction_shift(self, value):
        new_value = (value[0] * 1e-3, value[1] * 1e-3)
        self._scope.set("tem.Projection.DiffractionShift", new_value, vector=True)

    @property
    def diffraction_stigmator(self):
        """ Diffraction stigmator. (read/write)"""
        if self._scope.get("tem.Projection.Mode") == ProjectionMode.DIFFRACTION:
            return (self._scope.get("tem.Projection.DiffractionStigmator.X"),
                    self._scope.get("tem.Projection.DiffractionStigmator.Y"))
        else:
            raise RuntimeError(self._err_msg)

    @diffraction_stigmator.setter
    def diffraction_stigmator(self, value):
        if self._scope.get("tem.Projection.Mode") == ProjectionMode.DIFFRACTION:
            self._scope.set("tem.Projection.DiffractionStigmator", value, 
                           vector=True, limits=(-1.0, 1.0))
        else:
            raise RuntimeError(self._err_msg)

    @property
    def objective_stigmator(self):
        """ Objective stigmator. (read/write)"""
        return (self._scope.get("tem.Projection.ObjectiveStigmator.X"),
                self._scope.get("tem.Projection.ObjectiveStigmator.Y"))

    @objective_stigmator.setter
    def objective_stigmator(self, value):
        self._scope.set("tem.Projection.ObjectiveStigmator", value, 
                       vector=True, limits=(-1.0, 1.0))

    @property
    def defocus(self):
        """ Defocus value in um. (read/write)"""
        return self._scope.get("tem.Projection.Defocus") * 1e6

    @defocus.setter
    def defocus(self, value):
        self._scope.set("tem.Projection.Defocus", float(value) * 1e-6)

    @property
    def mode(self):
        """ Main mode of the projection system (either imaging or diffraction). (read/write)"""
        return ProjectionMode(self._scope.get("tem.Projection.Mode")).name

    @mode.setter
    def mode(self, mode):
        self._scope.set("tem.Projection.Mode", mode)

    @property
    def detector_shift(self):
        """ Detector shift. (read/write)"""
        return ProjectionDetectorShift(self._scope.get("tem.Projection.DetectorShift")).name

    @detector_shift.setter
    def detector_shift(self, value):
        self._scope.set("tem.Projection.DetectorShift", value)

    @property
    def detector_shift_mode(self):
        """ Detector shift mode. (read/write)"""
        return ProjDetectorShiftMode(self._scope.get("tem.Projection.DetectorShiftMode")).name

    @detector_shift_mode.setter
    def detector_shift_mode(self, value):
        self._scope.set("tem.Projection.DetectorShiftMode", value)

    @property
    def magnification_range(self):
        """ Submode of the projection system (either LM, M, SA, MH, LAD or D).
        The imaging submode can change when the magnification is changed.
        """
        return ProjectionSubMode(self._scope.get("tem.Projection.SubMode")).name

    @property
    def image_rotation(self):
        """ The rotation of the image or diffraction pattern on the
        fluorescent screen with respect to the specimen. Units: mrad.
        """
        return self._scope.get("tem.Projection.ImageRotation") * 1e3

    @property
    def is_eftem_on(self):
        """ Check if the EFTEM lens program setting is ON. """
        return LensProg(self._scope.get("tem.Projection.LensProgram")) == LensProg.EFTEM

    def eftem_on(self):
        """ Switch on EFTEM. """
        self._scope.set("tem.Projection.LensProgram", LensProg.EFTEM)

    def eftem_off(self):
        """ Switch off EFTEM. """
        self._scope.set("tem.Projection.LensProgram", LensProg.REGULAR)

    def reset_defocus(self):
        """ Reset defocus value in the TEM user interface to zero.
        Does not change any lenses. """
        self._scope.exec("tem.Projection.ResetDefocus()")


class Apertures:
    """ Apertures and VPP controls. """
    def __init__(self, microscope):
        self._scope = microscope
        self._err_msg = "Apertures interface is not available. Requires a separate license"
        self._err_msg_vpp = "Either no VPP found or it's not enabled and inserted"
        try:
            self._tem_apertures = self._scope.get("tem.ApertureMechanismCollection")
        except:
            self._tem_apertures = None
            logging.info(self._err_msg)

    def _find_aperture(self, name):
        """Find aperture object by name. """
        if self._tem_apertures is None:
            raise NotImplementedError(self._err_msg)
        for ap in self._tem_apertures:
            if MechanismId(ap.Id).name == name.upper():
                return ap
        raise KeyError("No aperture with name %s" % name)

    @property
    def vpp_position(self):
        """ Returns the index of the current VPP preset position. """
        try:
            return self._scope.get("tem_adv.PhasePlate.GetCurrentPresetPosition") + 1
        except:
            raise RuntimeError(self._err_msg_vpp)

    def vpp_next_position(self):
        """ Goes to the next preset location on the VPP aperture. """
        try:
            self._scope.exec("tem_adv.PhasePlate.SelectNextPresetPosition()")
        except:
            raise RuntimeError(self._err_msg_vpp)

    def enable(self, aperture):
        ap = self._find_aperture(aperture)
        ap.Enable()

    def disable(self, aperture):
        ap = self._find_aperture(aperture)
        ap.Disable()

    def retract(self, aperture):
        ap = self._find_aperture(aperture)
        if ap.IsRetractable:
            ap.Retract()

    def select(self, aperture, size):
        """ Select a specific aperture.

        :param aperture: Aperture name (C1, C2, C3, OBJ or SA)
        :type aperture: str
        :param size: Aperture size
        :type size: float
        """
        ap = self._find_aperture(aperture)
        if ap.State == MechanismState.DISABLED:
            ap.Enable()
        for a in ap.ApertureCollection:
            if a.Diameter == size:
                ap.SelectAperture(a)
                if ap.SelectedAperture.Diameter == size:
                    return
                else:
                    raise RuntimeError("Could not select aperture!")

    @property
    def show_all(self):
        """ Returns a dict with apertures information. """
        if self._tem_apertures is None:
            raise NotImplementedError(self._err_msg)
        result = {}
        for ap in self._tem_apertures:
            result[MechanismId(ap.Id).name] = {"retractable": ap.IsRetractable,
                                               "state": MechanismState(ap.State).name,
                                               "sizes": [a.Diameter for a in ap.ApertureCollection]
                                               }
        return result


class Gun:
    """ Gun functions. """
    def __init__(self, microscope):
        self._scope = microscope
        self._tem_gun1 = self._scope.has("tem.Gun1")
        self._err_msg_gun1 = "Gun1 interface is not available"
        self._err_msg_cfeg = "Source/C-FEG interface is not available"

        if not self._tem_gun1:
            logging.info(self._err_msg_gun1 + ". Requires TEM Server 7.10+")
        try:
            self._tem_feg = self._scope.has("tem_adv.Source")
            _ = self._scope.get("tem_adv.Source.State")
        except:
            self._tem_feg = False
            logging.info(self._err_msg_cfeg)

    @property
    def shift(self):
        """ Gun shift. (read/write)"""
        return (self._scope.get("tem.Gun.Shift.X"),
                self._scope.get("tem.Gun.Shift.Y"))

    @shift.setter
    def shift(self, value):
        self._scope.set("tem.Gun.Shift", value, vector=True, limits=(-1.0, 1.0))

    @property
    def tilt(self):
        """ Gun tilt. (read/write)"""
        return (self._scope.get("tem.Gun.Tilt.X"),
                self._scope.get("tem.Gun.Tilt.Y"))

    @tilt.setter
    def tilt(self, value):
        self._scope.set("tem.Gun.Tilt", value, vector=True, limits=(-1.0, 1.0))

    @property
    def voltage_offset(self):
        """ High voltage offset. (read/write)"""
        if self._tem_gun1:
            return self._scope.get("tem.Gun1.HighVoltageOffset")
        else:
            raise NotImplementedError(self._err_msg_gun1)

    @voltage_offset.setter
    def voltage_offset(self, offset):
        if self._tem_gun1:
            self._scope.set("tem.Gun1.HighVoltageOffset", offset)
        else:
            raise NotImplementedError(self._err_msg_gun1)

    @property
    def feg_state(self):
        """ FEG emitter status. """
        if self._tem_feg:
            return FegState(self._scope.get("tem_adv.Source.State")).name
        else:
            raise NotImplementedError(self._err_msg_cfeg)

    @property
    def ht_state(self):
        """ High tension state: on, off or disabled.
        Disabling/enabling can only be done via the button on the
        system on/off-panel, not via script. When switching on
        the high tension, this function cannot check if and
        when the set value is actually reached. (read/write)
        """
        return HighTensionState(self._scope.get("tem.Gun.HTState")).name

    @ht_state.setter
    def ht_state(self, value):
        self._scope.set("tem.Gun.HTState", value)

    @property
    def voltage(self):
        """ The value of the HT setting as displayed in the TEM user
        interface. Units: kVolts. (read/write)
        """
        state = self._scope.get("tem.Gun.HTState")
        if state == HighTensionState.ON:
            return self._scope.get("tem.Gun.HTValue") * 1e-3
        else:
            return 0.0

    @voltage.setter
    def voltage(self, value):
        voltage_max = self.voltage_max
        if not (0.0 <= value <= voltage_max):
            raise ValueError("%s is outside of range 0.0-%s" % (value, voltage_max))
        self._scope.set("tem.Gun.HTValue", float(value) * 1000)
        while True:
            if self._scope.get("tem.Gun.HTValue") == float(value) * 1000:
                logging.info("Changing HT voltage complete.")
                break
            else:
                time.sleep(10)

    @property
    def voltage_max(self):
        """ The maximum possible value of the HT on this microscope. Units: kVolts. """
        return self._scope.get("tem.Gun.HTMaxValue") * 1e-3

    @property
    def voltage_offset_range(self):
        """ Returns the high voltage offset range. """
        if self._tem_gun1:
            #TODO: this is a function?
            return self._scope.exec("tem.Gun1.GetHighVoltageOffsetRange()")
        else:
            raise NotImplementedError(self._err_msg_gun1)

    @property
    def beam_current(self):
        """ Returns the C-FEG beam current in Amperes. """
        if self._tem_feg:
            return self._scope.get("tem_adv.Source.BeamCurrent")
        else:
            raise NotImplementedError(self._err_msg_cfeg)

    @property
    def extractor_voltage(self):
        """ Returns the extractor voltage. """
        if self._tem_feg:
            return self._scope.get("tem_adv.Source.ExtractorVoltage")
        else:
            raise NotImplementedError(self._err_msg_cfeg)

    @property
    def focus_index(self):
        """ Returns coarse and fine gun lens index. """
        if self._tem_feg:
            return (self._scope.get("tem_adv.Source.FocusIndex.Coarse"),
                    self._scope.get("tem_adv.Source.FocusIndex.Fine"))
        else:
            raise NotImplementedError(self._err_msg_cfeg)

    def do_flashing(self, flash_type):
        """ Perform cold FEG flashing.

        :param flash_type: FEG flashing type (FegFlashingType enum)
        :type flash_type: IntEnum
        """
        if not self._tem_feg:
            raise NotImplementedError(self._err_msg_cfeg)
        if self._scope.exec("tem_adv.Source.Flashing.IsFlashingAdvised()", flash_type):
            # FIXME: lowT flashing can be done even if not advised
            self._scope.exec("tem_adv.Source.Flashing.PerformFlashing()", flash_type)
        else:
            raise Warning("Flashing type %s is not advised" % flash_type)


class EnergyFilter:
    """ Energy filter controls. """
    def __init__(self, microscope):
        self._scope = microscope
        self._err_msg = "EnergyFilter interface is not available"
        self._tem_ef = self._scope.has("tem_adv.EnergyFilter")
        if not self._tem_ef:
            logging.info(self._err_msg)

    def _check_range(self, attrname, value):
        vmin = self._scope.get(attrname + ".Begin")
        vmax = self._scope.get(attrname + ".End")
        if not (vmin <= value <= vmax):
            raise ValueError("Value is outside of allowed "
                             "range: %0.3f - %0.3f" % (vmin, vmax))

    def insert_slit(self, width):
        """ Insert energy slit.

        :param width: Slit width in eV
        :type width: float
        """
        if not self._tem_ef:
            raise NotImplementedError(self._err_msg)
        self._check_range("tem_adv.EnergyFilter.Slit.WidthRange", width)
        self._scope.set("tem_adv.EnergyFilter.Slit.Width", width)
        if not self._scope.get("tem_adv.EnergyFilter.Slit.IsInserted"):
            self._scope.exec("tem_adv.EnergyFilter.Slit.Insert()")

    def retract_slit(self):
        """ Retract energy slit. """
        if not self._tem_ef:
            raise NotImplementedError(self._err_msg)
        self._scope.exec("tem_adv.EnergyFilter.Slit.Retract()")

    @property
    def slit_width(self):
        """ Returns energy slit width in eV. """
        if not self._tem_ef:
            raise NotImplementedError(self._err_msg)
        return self._scope.get("tem_adv.EnergyFilter.Slit.Width")

    @slit_width.setter
    def slit_width(self, value):
        if not self._tem_ef:
            raise NotImplementedError(self._err_msg)
        self._check_range("tem_adv.EnergyFilter.Slit.WidthRange", value)
        self._scope.set("tem_adv.EnergyFilter.Slit.Width", value)

    @property
    def ht_shift(self):
        """ Returns High Tension energy shift in eV. """
        if not self._tem_ef:
            raise NotImplementedError(self._err_msg)
        return self._scope.get("tem_adv.EnergyFilter.HighTensionEnergyShift.EnergyShift")

    @ht_shift.setter
    def ht_shift(self, value):
        if not self._tem_ef:
            raise NotImplementedError(self._err_msg)
        self._check_range("tem_adv.EnergyFilter.HighTensionEnergyShift.EnergyShiftRange", value)
        self._scope.set("tem_adv.EnergyFilter.HighTensionEnergyShift.EnergyShift", value)

    @property
    def zlp_shift(self):
        if not self._tem_ef:
            raise NotImplementedError(self._err_msg)
        """ Returns Zero-Loss Peak (ZLP) energy shift in eV. """
        return self._scope.get("tem_adv.EnergyFilter.ZeroLossPeakAdjustment.EnergyShift")

    @zlp_shift.setter
    def zlp_shift(self, value):
        if not self._tem_ef:
            raise NotImplementedError(self._err_msg)
        self._check_range("tem_adv.EnergyFilter.ZeroLossPeakAdjustment.EnergyShiftRange", value)
        self._scope.set("tem_adv.EnergyFilter.ZeroLossPeakAdjustment.EnergyShift", value)


class LowDose:
    """ Low Dose functions. """
    def __init__(self, microscope, skip_check=False):
        self._scope = microscope
        self._err_msg = "Low Dose is not available"
        if not skip_check and not self._scope.has("tem_lowdose"):
            logging.info(self._err_msg)

    @property
    def is_available(self):
        """ Return True if Low Dose is available. """
        return self._scope.get("tem_lowdose.LowDoseAvailable") and self._scope.get("tem_lowdose.IsInitialized")

    @property
    def is_active(self):
        """ Check if the Low Dose is ON. """
        if self.is_available:
            return LDStatus(self._scope.get("tem_lowdose.LowDoseActive")) == LDStatus.IS_ON
        else:
            raise RuntimeError(self._err_msg)

    @property
    def state(self):
        """ Low Dose state (LDState enum). (read/write) """
        if self.is_available and self.is_active:
            return LDState(self._scope.get("tem_lowdose.LowDoseState")).name
        else:
            raise RuntimeError(self._err_msg)

    @state.setter
    def state(self, state):
        if self.is_available:
            self._scope.set("tem_lowdose.LowDoseState", state)
        else:
            raise RuntimeError(self._err_msg)

    def on(self):
        """ Switch ON Low Dose."""
        if self.is_available:
            self._scope.set("tem_lowdose.LowDoseActive", LDStatus.IS_ON)
        else:
            raise RuntimeError(self._err_msg)

    def off(self):
        """ Switch OFF Low Dose."""
        if self.is_available:
            self._scope.set("tem_lowdose.LowDoseActive", LDStatus.IS_OFF)
        else:
            raise RuntimeError(self._err_msg)


class Image(BaseImage):
    """ Acquired image object. """
    def __init__(self, obj, name=None, isAdvanced=False, **kwargs):
        super().__init__(obj, name, isAdvanced, **kwargs)

    def _get_metadata(self, obj):
        return {item.Key: item.ValueAsString for item in obj.Metadata}

    @property
    def width(self):
        """ Image width in pixels. """
        return self._img.Width

    @property
    def height(self):
        """ Image height in pixels. """
        return self._img.Height

    @property
    def bit_depth(self):
        """ Bit depth. """
        return self._img.BitDepth if self._isAdvanced else self._img.Depth

    @property
    def pixel_type(self):
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

    def save(self, filename, normalize=False):
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
                    fmt = AcqImageFileFormat[fmt].value
                except KeyError:
                    raise NotImplementedError("Format %s is not supported" % fmt)
                self._img.AsFile(filename, fmt, normalize)
