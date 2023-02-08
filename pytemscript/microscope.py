import logging
import math
import time
import os
from datetime import datetime

from .utils.enums import *
from .base_microscope import BaseMicroscope, BaseImage, Vector


class Microscope(BaseMicroscope):
    """ High level interface to the local microscope.
    Creating an instance of this class already queries COM interfaces for the instrument.

    :param useLD: Connect to LowDose server on microscope PC (limited control only)
    :type useLD: bool
    :param useTecnaiCCD: Connect to TecnaiCCD plugin on microscope PC that controls Digital Micrograph (may be faster than via TIA / std scripting)
    :type useTecnaiCCD: bool
    :param useSEMCCD: Connect to SerialEMCCD plugin on Gatan PC that controls Digital Micrograph (may be faster than via TIA / std scripting)
    :type useSEMCCD: bool
    """
    def __init__(self, useLD=True, useTecnaiCCD=False, useSEMCCD=False, remote=False):

        super().__init__(useLD, useTecnaiCCD, useSEMCCD, remote)

        if useTecnaiCCD:
            if self._tecnai_ccd is None:
                raise RuntimeError("Could not use Tecnai CCD plugin, "
                                   "please set useTecnaiCCD=False")
            else:
                from .tecnai_ccd_plugin import TecnaiCCDPlugin
                self._tecnai_ccd_plugin = TecnaiCCDPlugin(self)

        if useSEMCCD:
            if self._sem_ccd is None:
                raise RuntimeError("Could not use SerialEM CCD plugin, "
                                   "please set useSEMCCD=False")
            else:
                from .serialem_ccd_plugin import SerialEMCCDPlugin
                self._sem_ccd_plugin = SerialEMCCDPlugin(self)

        self.acquisition = Acquisition(self)
        self.detectors = Detectors(self)
        self.gun = Gun(self)
        self.lowdose = LowDose(self)
        self.optics = Optics(self)
        self.stem = Stem(self)
        self.temperature = Temperature(self)
        self.vacuum = Vacuum(self)
        self.autoloader = Autoloader(self)
        self.stage = Stage(self)
        self.piezo_stage = PiezoStage(self)
        self.apertures = Apertures(self)

        if self._tem_adv is not None:
            self.user_door = UserDoor(self)
            self.energy_filter = EnergyFilter(self)

        if useLD:
            self.lowdose = LowDose(self)

    @property
    def family(self):
        """ Returns the microscope product family / platform. """
        return ProductFamily(self._tem.Configuration.ProductFamily).name

    @property
    def condenser_system(self):
        """ Returns the type of condenser lens system: two or three lenses. """
        return CondenserLensSystem(self._tem.Configuration.CondenserLensSystem).name

    @property
    def user_buttons(self):
        """ Returns a dict with assigned hand panels buttons. """
        return {b.Name: b.Label for b in self._tem.UserButtons}


class UserDoor:
    """ User door hatch controls. """
    def __init__(self, microscope):
        if hasattr(microscope._tem_adv, "UserDoorHatch"):
            self._tem_door = microscope._tem_adv.UserDoorHatch
        else:
            self._tem_door = None

    @property
    def state(self):
        """ Returns door state. """
        return HatchState(self._tem_door.State).name

    def open(self):
        """ Open the door. """
        if self._tem_door.IsControlAllowed:
            self._tem_door.Open()
        else:
            raise RuntimeError("Door control is unavailable")

    def close(self):
        """ Close the door. """
        if self._tem_door.IsControlAllowed:
            self._tem_door.Close()
        else:
            raise RuntimeError("Door control is unavailable")


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
        self._tem_temp_control = microscope._tem.TemperatureControl

        if microscope._tem_adv is not None and hasattr(microscope._tem_adv, "TemperatureControl"):
            self._tem_temp_control_adv = microscope._tem_adv.TemperatureControl
        else:
            self._tem_temp_control_adv = None

    @property
    def is_available(self):
        """ Status of the temperature control. Should be always False on Tecnai instruments. """
        return self._tem_temp_control.TemperatureControlAvailable

    def force_refill(self):
        """ Forces LN refill if the level is below 70%, otherwise returns an error.
        Note: this function takes considerable time to execute.
        """
        if self.is_available:
            self._tem_temp_control.ForceRefill()
        elif self._tem_temp_control_adv is not None:
            return self._tem_temp_control_adv.RefillAllDewars()
        else:
            raise RuntimeError("TemperatureControl is not available")

    def dewar_level(self, dewar):
        """ Returns the LN level (%) in a dewar.

        :param dewar: Dewar name (RefrigerantDewar enum)
        :type dewar: IntEnum
        """
        if self.is_available:
            return self._tem_temp_control.RefrigerantLevel(dewar)
        else:
            raise RuntimeError("TemperatureControl is not available")

    @property
    def is_dewar_filling(self):
        """ Returns TRUE if any of the dewars is currently busy filling. """
        if self.is_available:
            return self._tem_temp_control.DewarsAreBusyFilling
        elif self._tem_temp_control_adv is not None:
            return self._tem_temp_control_adv.IsAnyDewarFilling
        else:
            raise RuntimeError("TemperatureControl is not available")

    @property
    def dewars_time(self):
        """ Returns remaining time (seconds) until the next dewar refill.
        Returns -1 if no refill is scheduled (e.g. All room temperature, or no
        dewar present).
        """
        # TODO: check if returns -60 at room temperature
        if self.is_available:
            return self._tem_temp_control.DewarsRemainingTime
        else:
            raise RuntimeError("TemperatureControl is not available")

    @property
    def temp_docker(self):
        """ Returns Docker temperature in Kelvins. """
        if self._tem_temp_control_adv is not None:
            return self._tem_temp_control_adv.AutoloaderCompartment.DockerTemperature
        else:
            raise NotImplementedError("This function is not available "
                                      "in your adv. scripting interface.")

    @property
    def temp_cassette(self):
        """ Returns Cassette gripper temperature in Kelvins. """
        if self._tem_temp_control_adv is not None:
            return self._tem_temp_control_adv.AutoloaderCompartment.CassetteTemperature
        else:
            raise NotImplementedError("This function is not available "
                                      "in your adv. scripting interface.")

    @property
    def temp_cartridge(self):
        """ Returns Cartridge gripper temperature in Kelvins. """
        if self._tem_temp_control_adv is not None:
            return self._tem_temp_control_adv.AutoloaderCompartment.CartridgeTemperature
        else:
            raise NotImplementedError("This function is not available "
                                      "in your adv. scripting interface.")

    @property
    def temp_holder(self):
        """ Returns Holder temperature in Kelvins. """
        if self._tem_temp_control_adv is not None:
            return self._tem_temp_control_adv.ColumnCompartment.HolderTemperature
        else:
            raise NotImplementedError("This function is not available "
                                      "in your adv. scripting interface.")


class Autoloader:
    """ Sample loading functions. """
    def __init__(self, microscope):
        self._tem_autoloader = microscope._tem.AutoLoader

        if microscope._tem_adv is not None and hasattr(microscope._tem_adv, "AutoLoader"):
            self._tem_autoloader_adv = microscope._tem_adv.AutoLoader
        else:
            self._tem_autoloader_adv = None

    @property
    def is_available(self):
        """ Status of the autoloader. Should be always False on Tecnai instruments. """
        return self._tem_autoloader.AutoLoaderAvailable

    @property
    def number_of_slots(self):
        """ The number of slots in a cassette. """
        if self.is_available:
            return self._tem_autoloader.NumberOfCassetteSlots
        else:
            raise RuntimeError("Autoloader is not available")

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
            self._tem_autoloader.LoadCartridge(slot)
        else:
            raise RuntimeError("Autoloader is not available")

    def unload_cartridge(self):
        """ Unloads the cartridge currently in the microscope and puts it back into its
        slot in the cassette.
        """
        if self.is_available:
            self._tem_autoloader.UnloadCartridge()
        else:
            raise RuntimeError("Autoloader is not available")

    def run_inventory(self):
        """ Performs an inventory of the cassette.
        Note: This function takes considerable time to execute.
        """
        # TODO: check if cassette is present
        if self.is_available:
            self._tem_autoloader.PerformCassetteInventory()
        else:
            raise RuntimeError("Autoloader is not available")

    def slot_status(self, slot):
        """ The status of the slot specified.

        :param slot: Slot number
        :type slot: int
        """
        if self.is_available:
            total = self.number_of_slots
            if slot > total:
                raise ValueError("Only %s slots are available" % total)
            status = self._tem_autoloader.SlotStatus(int(slot))
            return CassetteSlotStatus(status).name
        else:
            raise RuntimeError("Autoloader is not available")

    def undock_cassette(self):
        """ Moves the cassette from the docker to the capsule. """
        if self._tem_autoloader_adv is not None:
            if self.is_available:
                self._tem_autoloader_adv.UndockCassette()
            else:
                raise RuntimeError("Autoloader is not available")
        else:
            raise NotImplementedError("This function is not available "
                                      "in your adv. scripting interface.")

    def dock_cassette(self):
        """ Moves the cassette from the capsule to the docker. """
        if self._tem_autoloader_adv is not None:
            if self.is_available:
                self._tem_autoloader_adv.DockCassette()
            else:
                raise RuntimeError("Autoloader is not available")
        else:
            raise NotImplementedError("This function is not available "
                                      "in your adv. scripting interface.")

    def initialize(self):
        """ Initializes / Recovers the Autoloader for further use. """
        if self._tem_autoloader_adv is not None:
            if self.is_available:
                self._tem_autoloader_adv.Initialize()
            else:
                raise RuntimeError("Autoloader is not available")
        else:
            raise NotImplementedError("This function is not available "
                                      "in your adv. scripting interface.")

    def buffer_cycle(self):
        """ Synchronously runs the Autoloader buffer cycle. """
        if self._tem_autoloader_adv is not None:
            if self.is_available:
                self._tem_autoloader_adv.BufferCycle()
            else:
                raise RuntimeError("Autoloader is not available")
        else:
            raise NotImplementedError("This function is not available "
                                      "in your adv. scripting interface.")


class Stage:
    """ Stage functions. """
    def __init__(self, microscope):
        self._tem_stage = microscope._tem.Stage

    def _from_dict(self, **values):
        axes = 0
        position = self._tem_stage.Position
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
            if self._tem_stage.Status != StageStatus.READY:
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
                    self._tem_stage.MoveTo(new_pos, axes)
                else:
                    if speed is not None:
                        self._tem_stage.GoToWithSpeed(new_pos, axes, speed)
                    else:
                        self._tem_stage.GoTo(new_pos, axes)

                break
        else:
            raise RuntimeError("Stage is not ready.")

    @property
    def status(self):
        """ The current state of the stage. """
        return StageStatus(self._tem_stage.Status).name

    @property
    def holder(self):
        """ The current specimen holder type. """
        return StageHolderType(self._tem_stage.Holder).name

    @property
    def position(self):
        """ The current position of the stage (x,y,z in um and a,b in degrees). """
        pos = self._tem_stage.Position
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
            data = self._tem_stage.AxisData(StageAxes[axis.upper()])
            result[axis] = {
                'min': data.MinPos,
                'max': data.MaxPos,
                'unit': MeasurementUnitType(data.UnitType).name
            }
        return result


class PiezoStage:
    """ Piezo stage functions. """
    def __init__(self, microscope):
        try:
            self._tem_pstage = microscope._tem_adv.PiezoStage
            self.high_resolution = self._tem_pstage.HighResolution
        except:
            logging.info("PiezoStage interface is not available.")

    @property
    def position(self):
        """ The current position of the piezo stage (x,y,z in um). """
        pos = self._tem_pstage.CurrentPosition
        return {key: getattr(pos, key.upper()) * 1e6 for key in 'xyz'}

    @property
    def position_range(self):
        """ Return min and max positions. """
        return self._tem_pstage.GetPositionRange()

    @property
    def velocity(self):
        """ Returns a dict with stage velocities. """
        pos = self._tem_pstage.CurrentJogVelocity
        return {key: getattr(pos, key.upper()) for key in 'xyz'}


class Vacuum:
    """ Vacuum functions. """
    def __init__(self, microscope):
        self._tem_vacuum = microscope._tem.Vacuum

    @property
    def status(self):
        """ Status of the vacuum system. """
        return VacuumStatus(self._tem_vacuum.Status).name

    @property
    def is_buffer_running(self):
        """ Checks whether the prevacuum pump is currently running
        (consequences: vibrations, exposure function blocked
        or should not be called).
        """
        return self._tem_vacuum.PVPRunning

    @property
    def is_column_open(self):
        """ The status of the column valves. """
        return self._tem_vacuum.ColumnValvesOpen

    @property
    def gauges(self):
        """ Returns a dict with vacuum gauges information.
        Pressure values are in Pascals.
        """
        gauges = {}
        for g in self._tem_vacuum.Gauges:
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
        self._tem_vacuum.ColumnValvesOpen = True

    def column_close(self):
        """ Close column valves. """
        self._tem_vacuum.ColumnValvesOpen = False

    def run_buffer_cycle(self):
        """ Runs a pumping cycle to empty the buffer. """
        self._tem_vacuum.RunBufferCycle()


class Optics:
    """ Projection, Illumination functions. """
    def __init__(self, microscope):
        self._tem = microscope._tem
        self._tem_adv = microscope._tem_adv
        self._tem_cam = self._tem.Camera
        self._tem_illumination = self._tem.Illumination
        self._tem_projection = self._tem.Projection
        self._tem_control = self._tem.InstrumentModeControl

        self.illumination = Illumination(self._tem)
        self.projection = Projection(self._tem_projection)

    @property
    def screen_current(self):
        """ The current measured on the fluorescent screen (units: nanoAmperes). """
        return self._tem_cam.ScreenCurrent * 1e9

    @property
    def is_beam_blanked(self):
        """ Status of the beam blanker. """
        return self._tem_illumination.BeamBlanked

    @property
    def is_shutter_override_on(self):
        """ Determines the state of the shutter override function.
        WARNING: Do not leave the Shutter override on when stopping the script.
        The microscope operator will be unable to have a beam come down and has
        no separate way of seeing that it is blocked by the closed microscope shutter.
        """
        return self._tem.BlankerShutter.ShutterOverrideOn

    @property
    def is_autonormalize_on(self):
        """ Status of the automatic normalization procedures performed by
        the TEM microscope. Normally they are active, but for scripting it can be
        convenient to disable them temporarily.
        """
        return self._tem.AutoNormalizeEnabled

    def beam_blank(self):
        """ Activates the beam blanker. """
        self._tem_illumination.BeamBlanked = True
        logging.warning("Falcon protector might delay blanker response")

    def beam_unblank(self):
        """ Deactivates the beam blanker. """
        self._tem_illumination.BeamBlanked = False
        logging.warning("Falcon protector might delay blanker response")

    def normalize_all(self):
        """ Normalize all lenses. """
        self._tem.NormalizeAll()

    def normalize(self, mode):
        """ Normalize condenser or projection lens system.
        :param mode: Normalization mode (ProjectionNormalization or IlluminationNormalization enum)
        :type mode: IntEnum
        """
        if mode in ProjectionNormalization:
            self._tem_projection.Normalize(mode)
        elif mode in IlluminationNormalization:
            self._tem_illumination.Normalize(mode)
        else:
            raise ValueError("Unknown normalization mode: %s" % mode)


class Stem:
    """ STEM functions. """
    def __init__(self, microscope):
        self._tem = microscope._tem
        self._tem_illumination = self._tem.Illumination
        self._tem_control = self._tem.InstrumentModeControl

    @property
    def is_available(self):
        """ Returns whether the microscope has a STEM system or not. """
        return self._tem_control.StemAvailable

    def enable(self):
        """ Switch to STEM mode."""
        if self.is_available:
            self._tem_control.InstrumentMode = InstrumentMode.STEM
        else:
            raise RuntimeError("No STEM mode available")

    def disable(self):
        """ Switch back to TEM mode. """
        self._tem_control.InstrumentMode = InstrumentMode.TEM

    @property
    def magnification(self):
        """ The magnification value in STEM mode. (read/write)"""
        if self._tem_control.InstrumentMode == InstrumentMode.STEM:
            return self._tem_illumination.StemMagnification
        else:
            raise RuntimeError("Microscope not in STEM mode.")

    @magnification.setter
    def magnification(self, mag):
        if self._tem_control.InstrumentMode == InstrumentMode.STEM:
            self._tem_illumination.StemMagnification = float(mag)
        else:
            raise RuntimeError("Microscope not in STEM mode.")

    @property
    def rotation(self):
        """ The STEM rotation angle (in mrad). (read/write)"""
        if self._tem_control.InstrumentMode == InstrumentMode.STEM:
            return self._tem_illumination.StemRotation * 1e3
        else:
            raise RuntimeError("Microscope not in STEM mode.")

    @rotation.setter
    def rotation(self, rot):
        if self._tem_control.InstrumentMode == InstrumentMode.STEM:
            self._tem_illumination.StemRotation = float(rot) * 1e-3
        else:
            raise RuntimeError("Microscope not in STEM mode.")

    @property
    def scan_field_of_view(self):
        """ STEM full scan field of view. (read/write)"""
        if self._tem_control.InstrumentMode == InstrumentMode.STEM:
            return (self._tem_illumination.StemFullScanFieldOfView.X,
                    self._tem_illumination.StemFullScanFieldOfView.Y)
        else:
            raise RuntimeError("Microscope not in STEM mode.")

    @scan_field_of_view.setter
    def scan_field_of_view(self, values):
        if self._tem_control.InstrumentMode == InstrumentMode.STEM:
            Vector.set(self._tem_illumination, "StemFullScanFieldOfView", values)
        else:
            raise RuntimeError("Microscope not in STEM mode.")


class Illumination:
    """ Illumination functions. """
    def __init__(self, tem):
        self._tem = tem
        self._tem_illumination = self._tem.Illumination

    @property
    def spotsize(self):
        """ Spotsize number, usually 1 to 11. (read/write)"""
        return self._tem_illumination.SpotsizeIndex

    @spotsize.setter
    def spotsize(self, value):
        if not (1 <= int(value) <= 11):
            raise ValueError("%s is outside of range 1-11" % value)
        self._tem_illumination.SpotsizeIndex = int(value)

    @property
    def intensity(self):
        """ Intensity / C2 condenser lens value. (read/write)"""
        return self._tem_illumination.Intensity

    @intensity.setter
    def intensity(self, value):
        if not (0.0 <= value <= 1.0):
            raise ValueError("%s is outside of range 0.0-1.0" % value)
        self._tem_illumination.Intensity = float(value)

    @property
    def intensity_zoom(self):
        """ Intensity zoom. Set to False to disable. (read/write)"""
        return self._tem_illumination.IntensityZoomEnabled

    @intensity_zoom.setter
    def intensity_zoom(self, value):
        self._tem_illumination.IntensityZoomEnabled = bool(value)

    @property
    def intensity_limit(self):
        """ Intensity limit. Set to False to disable. (read/write)"""
        return self._tem_illumination.IntensityLimitEnabled

    @intensity_limit.setter
    def intensity_limit(self, value):
        self._tem_illumination.IntensityLimitEnabled = bool(value)

    @property
    def beam_shift(self):
        """ Beam shift X and Y in um. (read/write)"""
        return (self._tem_illumination.Shift.X * 1e6,
                self._tem_illumination.Shift.Y * 1e6)

    @beam_shift.setter
    def beam_shift(self, value):
        new_value = (value[0] * 1e-6, value[1] * 1e-6)
        Vector.set(self._tem_illumination, "Shift", new_value)

    @property
    def rotation_center(self):
        """ Rotation center X and Y in mrad. (read/write)
            Depending on the scripting version,
            the values might need scaling by 6.0 to get mrads.
        """
        return (self._tem_illumination.RotationCenter.X * 1e3,
                self._tem_illumination.RotationCenter.Y * 1e3)

    @rotation_center.setter
    def rotation_center(self, value):
        new_value = (value[0] * 1e-3, value[1] * 1e-3)
        Vector.set(self._tem_illumination, "RotationCenter", new_value)

    @property
    def condenser_stigmator(self):
        """ C2 condenser stigmator X and Y. (read/write)"""
        return (self._tem_illumination.CondenserStigmator.X,
                self._tem_illumination.CondenserStigmator.Y)

    @condenser_stigmator.setter
    def condenser_stigmator(self, value):
        Vector.set(self._tem_illumination, "CondenserStigmator", value, range=(-1.0, 1.0))

    @property
    def illuminated_area(self):
        """ Illuminated area. Works only on 3-condenser lens systems. (read/write)"""
        if self._tem.Configuration.CondenserLensSystem == CondenserLensSystem.THREE_CONDENSER_LENSES:
            return self._tem_illumination.IlluminatedArea
        else:
            raise NotImplementedError("Illuminated area exists only on 3-condenser lens systems.")

    @illuminated_area.setter
    def illuminated_area(self, value):
        if self._tem.Configuration.CondenserLensSystem == CondenserLensSystem.THREE_CONDENSER_LENSES:
            self._tem_illumination.IlluminatedArea = float(value)
        else:
            raise NotImplementedError("Illuminated area exists only on 3-condenser lens systems.")

    @property
    def probe_defocus(self):
        """ Probe defocus. Works only on 3-condenser lens systems. (read/write)"""
        if self._tem.Configuration.CondenserLensSystem == CondenserLensSystem.THREE_CONDENSER_LENSES:
            return self._tem_illumination.ProbeDefocus
        else:
            raise NotImplementedError("Probe defocus exists only on 3-condenser lens systems.")

    @probe_defocus.setter
    def probe_defocus(self, value):
        if self._tem.Configuration.CondenserLensSystem == CondenserLensSystem.THREE_CONDENSER_LENSES:
            self._tem_illumination.ProbeDefocus = float(value)
        else:
            raise NotImplementedError("Probe defocus exists only on 3-condenser lens systems.")

    #TODO: check if the illum. mode is probe?
    @property
    def convergence_angle(self):
        """ Convergence angle. Works only on 3-condenser lens systems. (read/write)"""
        if self._tem.Configuration.CondenserLensSystem == CondenserLensSystem.THREE_CONDENSER_LENSES:
            return self._tem_illumination.ConvergenceAngle
        else:
            raise NotImplementedError("Convergence angle exists only on 3-condenser lens systems.")

    @convergence_angle.setter
    def convergence_angle(self, value):
        if self._tem.Configuration.CondenserLensSystem == CondenserLensSystem.THREE_CONDENSER_LENSES:
            self._tem_illumination.ConvergenceAngle = float(value)
        else:
            raise NotImplementedError("Convergence angle exists only on 3-condenser lens systems.")

    @property
    def C3ImageDistanceParallelOffset(self):
        """ C3 image distance parallel offset. Works only on 3-condenser lens systems. (read/write)"""
        if self._tem.Configuration.CondenserLensSystem == CondenserLensSystem.THREE_CONDENSER_LENSES:
            return self._tem_illumination.C3ImageDistanceParallelOffset
        else:
            raise NotImplementedError("C3ImageDistanceParallelOffset exists only on 3-condenser lens systems.")

    @C3ImageDistanceParallelOffset.setter
    def C3ImageDistanceParallelOffset(self, value):
        if self._tem.Configuration.CondenserLensSystem == CondenserLensSystem.THREE_CONDENSER_LENSES:
            self._tem_illumination.C3ImageDistanceParallelOffset = float(value)
        else:
            raise NotImplementedError("C3ImageDistanceParallelOffset exists only on 3-condenser lens systems.")

    @property
    def mode(self):
        """ Illumination mode: microprobe or nanoprobe. (read/write)"""
        return IlluminationMode(self._tem_illumination.Mode).name

    @mode.setter
    def mode(self, value):
        self._tem_illumination.Mode = value

    @property
    def dark_field(self):
        """ Dark field mode: cartesian, conical or off. (read/write)"""
        return DarkFieldMode(self._tem_illumination.DFMode).name

    @dark_field.setter
    def dark_field(self, value):
        self._tem_illumination.DFMode = value

    @property
    def condenser_mode(self):
        """ Mode of the illumination system: parallel or probe. (read/write)"""
        if self._tem.Configuration.CondenserLensSystem == CondenserLensSystem.THREE_CONDENSER_LENSES:
            return CondenserMode(self._tem_illumination.CondenserMode).name
        else:
            raise NotImplementedError("Condenser mode exists only on 3-condenser lens systems.")

    @condenser_mode.setter
    def condenser_mode(self, value):
        if self._tem.Configuration.CondenserLensSystem == CondenserLensSystem.THREE_CONDENSER_LENSES:
            self._tem_illumination.CondenserMode = value
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
        mode = self._tem_illumination.DFMode
        tilt = self._tem_illumination.Tilt
        if mode == DarkFieldMode.CONICAL:
            return tilt[0] * 1e3 * math.cos(tilt[1]), tilt[0] * 1e3 * math.sin(tilt[1])
        elif mode == DarkFieldMode.CARTESIAN:
            return tilt * 1e3
        else:
            return 0.0, 0.0  # Microscope might return nonsense if DFMode is OFF

    @beam_tilt.setter
    def beam_tilt(self, tilt):
        mode = self._tem_illumination.DFMode
        tilt[0] *= 1e-3
        tilt[1] *= 1e-3
        if tilt[0] == 0.0 and tilt[1] == 0.0:
            self._tem_illumination.Tilt = 0.0, 0.0
            self._tem_illumination.DFMode = DarkFieldMode.OFF
        elif mode == DarkFieldMode.CONICAL:
            self._tem_illumination.Tilt = math.sqrt(tilt[0] ** 2 + tilt[1] ** 2), math.atan2(tilt[1], tilt[0])
        elif mode == DarkFieldMode.OFF:
            self._tem_illumination.DFMode = DarkFieldMode.CARTESIAN
            self._tem_illumination.Tilt = tilt
        else:
            self._tem_illumination.Tilt = tilt


class Projection:
    """ Projection system functions. """
    def __init__(self, projection):
        self._tem_projection = projection
        #self.magnification_index = self._tem_projection.MagnificationIndex
        #self.camera_length_index = self._tem_projection.CameraLengthIndex

    @property
    def focus(self):
        """ Absolute focus value. (read/write)"""
        return self._tem_projection.Focus

    @focus.setter
    def focus(self, value):
        if not (-1.0 <= value <= 1.0):
            raise ValueError("%s is outside of range -1.0 to 1.0" % value)

        self._tem_projection.Focus = float(value)

    @property
    def magnification(self):
        """ The reference magnification value (screen up setting)."""
        if self._tem_projection.Mode == ProjectionMode.IMAGING:
            return self._tem_projection.Magnification
        else:
            raise RuntimeError("Microscope is in diffraction mode.")

    @property
    def camera_length(self):
        """ The reference camera length in m (screen up setting). """
        if self._tem_projection.Mode == ProjectionMode.DIFFRACTION:
            return self._tem_projection.CameraLength
        else:
            raise RuntimeError("Microscope is not in diffraction mode.")

    @property
    def image_shift(self):
        """ Image shift in um. (read/write)"""
        return (self._tem_projection.ImageShift.X * 1e6,
                self._tem_projection.ImageShift.Y * 1e6)

    @image_shift.setter
    def image_shift(self, value):
        new_value = (value[0] * 1e-6, value[1] * 1e-6)
        Vector.set(self._tem_projection, "ImageShift", new_value)

    @property
    def image_beam_shift(self):
        """ Image shift with beam shift compensation in um. (read/write)"""
        return (self._tem_projection.ImageBeamShift.X * 1e6,
                self._tem_projection.ImageBeamShift.Y * 1e6)

    @image_beam_shift.setter
    def image_beam_shift(self, value):
        new_value = (value[0] * 1e-6, value[1] * 1e-6)
        Vector.set(self._tem_projection, "ImageBeamShift", new_value)

    @property
    def image_beam_tilt(self):
        """ Beam tilt with diffraction shift compensation in mrad. (read/write)"""
        return (self._tem_projection.ImageBeamTilt.X * 1e3,
                self._tem_projection.ImageBeamTilt.Y * 1e3)

    @image_beam_tilt.setter
    def image_beam_tilt(self, value):
        new_value = (value[0] * 1e-3, value[1] * 1e-3)
        Vector.set(self._tem_projection, "ImageBeamTilt", new_value)

    @property
    def diffraction_shift(self):
        """ Diffraction shift in mrad. (read/write)"""
        #TODO: 180/pi*value = approx number in TUI
        return (self._tem_projection.DiffractionShift.X * 1e3,
                self._tem_projection.DiffractionShift.Y * 1e3)

    @diffraction_shift.setter
    def diffraction_shift(self, value):
        new_value = (value[0] * 1e-3, value[1] * 1e-3)
        Vector.set(self._tem_projection, "DiffractionShift", new_value)

    @property
    def diffraction_stigmator(self):
        """ Diffraction stigmator. (read/write)"""
        if self._tem_projection.Mode == ProjectionMode.DIFFRACTION:
            return (self._tem_projection.DiffractionStigmator.X,
                    self._tem_projection.DiffractionStigmator.Y)
        else:
            raise RuntimeError("Microscope is not in diffraction mode.")

    @diffraction_stigmator.setter
    def diffraction_stigmator(self, value):
        if self._tem_projection.Mode == ProjectionMode.DIFFRACTION:
            Vector.set(self._tem_projection, "DiffractionStigmator",
                       value, range=(-1.0, 1.0))
        else:
            raise RuntimeError("Microscope is not in diffraction mode.")

    @property
    def objective_stigmator(self):
        """ Objective stigmator. (read/write)"""
        return (self._tem_projection.ObjectiveStigmator.X,
                self._tem_projection.ObjectiveStigmator.Y)

    @objective_stigmator.setter
    def objective_stigmator(self, value):
        Vector.set(self._tem_projection, "ObjectiveStigmator",
                   value, range=(-1.0, 1.0))

    @property
    def defocus(self):
        """ Defocus value in um. (read/write)"""
        return self._tem_projection.Defocus * 1e6

    @defocus.setter
    def defocus(self, value):
        self._tem_projection.Defocus = float(value) * 1e-6

    @property
    def mode(self):
        """ Main mode of the projection system (either imaging or diffraction). (read/write)"""
        return ProjectionMode(self._tem_projection.Mode).name

    @mode.setter
    def mode(self, mode):
        self._tem_projection.Mode = mode

    @property
    def detector_shift(self):
        """ Detector shift. (read/write)"""
        return ProjectionDetectorShift(self._tem_projection.DetectorShift).name

    @detector_shift.setter
    def detector_shift(self, value):
        self._tem_projection.DetectorShift = value

    @property
    def detector_shift_mode(self):
        """ Detector shift mode. (read/write)"""
        return ProjDetectorShiftMode(self._tem_projection.DetectorShiftMode).name

    @detector_shift_mode.setter
    def detector_shift_mode(self, value):
        self._tem_projection.DetectorShiftMode = value

    @property
    def magnification_range(self):
        """ Submode of the projection system (either LM, M, SA, MH, LAD or D).
        The imaging submode can change when the magnification is changed.
        """
        return ProjectionSubMode(self._tem_projection.SubMode).name

    @property
    def image_rotation(self):
        """ The rotation of the image or diffraction pattern on the
        fluorescent screen with respect to the specimen. Units: mrad.
        """
        return self._tem_projection.ImageRotation * 1e3

    @property
    def is_eftem_on(self):
        """ Check if the EFTEM lens program setting is ON. """
        return LensProg(self._tem_projection.LensProgram) == LensProg.EFTEM

    def eftem_on(self):
        """ Switch on EFTEM. """
        self._tem_projection.LensProgram = LensProg.EFTEM

    def eftem_off(self):
        """ Switch off EFTEM. """
        self._tem_projection.LensProgram = LensProg.REGULAR

    def reset_defocus(self):
        """ Reset defocus value in the TEM user interface to zero.
        Does not change any lenses. """
        self._tem_projection.ResetDefocus()


class Apertures:
    """ Apertures and VPP controls. """
    def __init__(self, microscope):
        self._has_advanced = microscope._tem_adv is not None
        if self._has_advanced:
            self._tem_vpp = microscope._tem_adv.PhasePlate

        try:
            self._tem_apertures = microscope._tem.ApertureMechanismCollection
        except:
            self._tem_apertures = None
            logging.info("Apertures interface is not available. Requires a separate license")

    def _find_aperture(self, name):
        """Find aperture object by name. """
        if self._tem_apertures is None:
            raise NotImplementedError("Apertures interface is not available. Requires a separate license")
        for ap in self._tem_apertures:
            if MechanismId(ap.Id).name == name.upper():
                return ap
        raise KeyError("No aperture with name %s" % name)

    @property
    def vpp_position(self):
        """ Returns the index of the current VPP preset position. """
        try:
            return self._tem_vpp.GetCurrentPresetPosition + 1
        except:
            raise RuntimeError("Either no VPP found or it's not enabled and inserted.")

    def vpp_next_position(self):
        """ Goes to the next preset location on the VPP aperture. """
        try:
            self._tem_vpp.SelectNextPresetPosition()
        except:
            raise RuntimeError("Either no VPP found or it's not enabled and inserted.")

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
            raise NotImplementedError("Apertures interface is not available. "
                                      "Requires a separate license")
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
        self._tem_gun = microscope._tem.Gun
        self._tem_gun1 = None
        self._tem_feg = None

        if hasattr(microscope._tem, "Gun1"):
            self._tem_gun1 = microscope._tem.Gun1
        else:
            logging.info("Gun1 interface is not available. Requires TEM Server 7.10+")

        try:
            self._tem_feg = microscope._tem_adv.Source
            _ = self._tem_feg.State
        except:
            logging.info("Source/C-FEG interface is not available.")

    @property
    def shift(self):
        """ Gun shift. (read/write)"""
        return (self._tem_gun.Shift.X, self._tem_gun.Shift.Y)

    @shift.setter
    def shift(self, value):
        Vector.set(self._tem_gun, "Shift", value, range=(-1.0, 1.0))

    @property
    def tilt(self):
        """ Gun tilt. (read/write)"""
        return (self._tem_gun.Tilt.X, self._tem_gun.Tilt.Y)

    @tilt.setter
    def tilt(self, value):
        Vector.set(self._tem_gun, "Tilt", value, range=(-1.0, 1.0))

    @property
    def voltage_offset(self):
        """ High voltage offset. (read/write)"""
        if self._tem_gun1 is None:
            raise RuntimeError("Gun1 interface is not available.")
        return self._tem_gun1.HighVoltageOffset

    @voltage_offset.setter
    def voltage_offset(self, offset):
        if self._tem_gun1 is None:
            raise RuntimeError("Gun1 interface is not available.")
        self._tem_gun1.HighVoltageOffset = offset

    @property
    def feg_state(self):
        """ FEG emitter status. """
        if self._tem_feg is None:
            raise RuntimeError("Gun1 interface is not available.")
        return FegState(self._tem_feg.State).name

    @property
    def ht_state(self):
        """ High tension state: on, off or disabled.
        Disabling/enabling can only be done via the button on the
        system on/off-panel, not via script. When switching on
        the high tension, this function cannot check if and
        when the set value is actually reached. (read/write)
        """
        return HighTensionState(self._tem_gun.HTState).name

    @ht_state.setter
    def ht_state(self, value):
        self._tem_gun.State = value

    @property
    def voltage(self):
        """ The value of the HT setting as displayed in the TEM user
        interface. Units: kVolts. (read/write)
        """
        state = self._tem_gun.HTState
        if state == HighTensionState.ON:
            return self._tem_gun.HTValue * 1e-3
        else:
            return 0.0

    @voltage.setter
    def voltage(self, value):
        voltage_max = self.voltage_max
        if not (0.0 <= value <= voltage_max):
            raise ValueError("%s is outside of range 0.0-%s" % (value, voltage_max))
        self._tem_gun.HTValue = float(value) * 1000
        while True:
            if self._tem_gun.HTValue == float(value) * 1000:
                logging.info("Changing HT voltage complete.")
                break
            else:
                time.sleep(10)

    @property
    def voltage_max(self):
        """ The maximum possible value of the HT on this microscope. Units: kVolts. """
        return self._tem_gun.HTMaxValue * 1e-3

    @property
    def voltage_offset_range(self):
        """ Returns the high voltage offset range. """
        if self._tem_gun1 is None:
            raise RuntimeError("Gun1 interface is not available.")
        return self._tem_gun1.GetHighVoltageOffsetRange()

    @property
    def beam_current(self):
        """ Returns the C-FEG beam current in Amperes. """
        if self._tem_feg is None:
            raise RuntimeError("Source/C-FEG interface is not available.")
        return self._tem_feg.BeamCurrent

    @property
    def extractor_voltage(self):
        """ Returns the extractor voltage. """
        if self._tem_feg is None:
            raise RuntimeError("Source/C-FEG interface is not available.")
        return self._tem_feg.ExtractorVoltage

    @property
    def focus_index(self):
        """ Returns coarse and fine gun lens index. """
        if self._tem_feg is None:
            raise RuntimeError("Source/C-FEG interface is not available.")
        focus_index = self._tem_feg.FocusIndex
        return (focus_index.Coarse, focus_index.Fine)

    def do_flashing(self, flash_type):
        """ Perform cold FEG flashing.

        :param flash_type: FEG flashing type (FegFlashingType enum)
        :type flash_type: IntEnum
        """
        if self._tem_feg is None:
            raise RuntimeError("Source/C-FEG interface is not available.")
        if self._tem_feg.Flashing.IsFlashingAdvised(flash_type):
            # FIXME: lowT flashing can be done even if not advised
            self._tem_feg.Flashing.PerformFlashing(flash_type)
        else:
            raise Warning("Flashing type %s is not advised" % flash_type)


class EnergyFilter:
    """ Energy filter controls. """
    def __init__(self, microscope):
        if hasattr(microscope._tem_adv, "EnergyFilter"):
            self._tem_ef = microscope._tem_adv.EnergyFilter
        else:
            logging.info("EnergyFilter interface is not available.")

    def _check_range(self, ev_range, value):
        if not (ev_range.Begin <= value <= ev_range.End):
            raise ValueError("Value is outside of allowed "
                             "range: %0.0f - %0.0f" % (ev_range.Begin,
                                                       ev_range.End))

    def insert_slit(self, width):
        """ Insert energy slit.

        :param width: Slit width in eV
        :type width: float
        """
        self._check_range(self._tem_ef.Slit.WidthRange, width)
        self._tem_ef.Slit.Width = width
        if not self._tem_ef.Slit.IsInserted:
            self._tem_ef.Slit.Insert()

    def retract_slit(self):
        """ Retract energy slit. """
        self._tem_ef.Slit.Retract()

    @property
    def slit_width(self):
        """ Returns energy slit width in eV. """
        return self._tem_ef.Slit.Width

    @slit_width.setter
    def slit_width(self, value):
        self._check_range(self._tem_ef.Slit.WidthRange, value)
        self._tem_ef.Slit.Width = value

    @property
    def ht_shift(self):
        """ Returns High Tension energy shift in eV. """
        return self._tem_ef.HighTensionEnergyShift.EnergyShift

    @ht_shift.setter
    def ht_shift(self, value):
        self._check_range(self._tem_ef.HighTensionEnergyShift.EnergyShiftRange, value)
        self._tem_ef.HighTensionEnergyShift.EnergyShift = value

    @property
    def zlp_shift(self):
        """ Returns Zero-Loss Peak (ZLP) energy shift in eV. """
        return self._tem_ef.ZeroLossPeakAdjustment.EnergyShift

    @zlp_shift.setter
    def zlp_shift(self, value):
        self._check_range(self._tem_ef.ZeroLossPeakAdjustment.EnergyShiftRange, value)
        self._tem_ef.ZeroLossPeakAdjustment.EnergyShift = value


class LowDose:
    """ Low Dose functions. """
    def __init__(self, microscope):
        if microscope._lowdose is not None:
            self._tem_ld = microscope._lowdose
        else:
            logging.info("LowDose server is not available.")

    @property
    def is_available(self):
        """ Return True if Low Dose is available. """
        return self._tem_ld.LowDoseAvailable and self._tem_ld.IsInitialized

    @property
    def is_active(self):
        """ Check if the Low Dose is ON. """
        if self.is_available:
            return LDStatus(self._tem_ld.LowDoseActive) == LDStatus.IS_ON
        else:
            raise RuntimeError("Low Dose is not available")

    @property
    def state(self):
        """ Low Dose state (LDState enum). (read/write) """
        if self.is_available and self.is_active:
            return LDState(self._tem_ld.LowDoseState).name
        else:
            raise RuntimeError("Low Dose is not available")

    @state.setter
    def state(self, state):
        if self.is_available:
            self._tem_ld.LowDoseState = state
        else:
            raise RuntimeError("Low Dose is not available")

    def on(self):
        """ Switch ON Low Dose."""
        if self.is_available:
            self._tem_ld.LowDoseActive = LDStatus.IS_ON
        else:
            raise RuntimeError("Low Dose is not available")

    def off(self):
        """ Switch OFF Low Dose."""
        if self.is_available:
            self._tem_ld.LowDoseActive = LDStatus.IS_OFF
        else:
            raise RuntimeError("Low Dose is not available")


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
