import math
import time
from .base_microscope import BaseMicroscope, Image, Vector
from .utils.enums import *


class Microscope(BaseMicroscope):
    """
    High level interface to local microscope.
    Creating an instance of this class already queries the COM interface for the instrument.
    """
    def __init__(self, address=None, timeout=None, simulate=False):
        super().__init__(address, timeout, simulate)

        self.acquisition = Acquisition(self)
        self.detectors = Detectors(self)
        self.gun = Gun(self)
        self.optics = Optics(self)
        self.stem = Stem(self)
        self.apertures = Apertures(self)
        self.temperature = Temperature(self)
        self.vacuum = Vacuum(self)
        self.autoloader = Autoloader(self)
        self.stage = Stage(self)
        self.piezo_stage = PiezoStage(self)
        self.user_door = UserDoor(self)

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

    def check_license(self):
        """ Returns a dict with advanced scripting license information. """
        return {
            'tem_scripting': self._lic_cam.IsTemScriptingLicensed,
            'dose_fractions': self._lic_cam.IsDoseFractionsLicensed,
            'electron_counting': self._lic_cam.IsElectronCountingLicensed
        }


class UserDoor:
    """ User door hatch controls. """
    def __init__(self, microscope):
        try:
            self._tem_door = microscope._tem_adv.UserDoorHatch
        except:
            print("UserDoor interface is not available")

    @property
    def state(self):
        return HatchState(self._tem_door.State).name

    def open(self):
        if self._tem_door.IsControlAllowed:
            self._tem_door.Open()

    def close(self):
        if self._tem_door.IsControlAllowed:
            self._tem_door.Close()


class Acquisition:
    """ Image acquisition functions. """
    def __init__(self, microscope):
        self._tem = microscope._tem
        self._tem_acq = self._tem._tem.Acquisition
        self._tem_csa = microscope._tem_adv.Acquisitions.CameraSingleAcquisition
        self._tem_cam = self._tem.Camera
        self._is_advanced = False
        self._prev_shutter_mode = None

    def _find_camera(self, name):
        """Find camera object by name. Check adv scripting first. """
        for cam in self._tem_csa.SupportedCameras:
            if cam.Name == name:
                self._is_advanced = True
                return cam
        for cam in self._tem_acq.Cameras:
            if cam.Info.Name == name:
                return cam
        raise KeyError("No camera with name %s" % name)

    def _find_stem_detector(self, name):
        """Find STEM detector object by name"""
        for stem in self._tem_acq.Detectors:
            if stem.Info.Name == name:
                return stem
        raise KeyError("No STEM detector with name %s" % name)

    def _check_binning(self, binning, camera, is_advanced=False):
        """ Check if input binning is in SupportedBinnings.

        :param binning: Input binning
        :type binning: int
        :param camera: Camera object
        :param is_advanced: Is this an advanced camera?
        :type is_advanced: bool
        :returns: Binning object
        """
        if is_advanced:
            param = self._tem_csa.CameraSettings.Capabilities
            for b in param.SupportedBinnings:
                if int(b.Width) == int(binning):
                    return b
        else:
            info = camera.Info
            for b in info.Binnings:
                if int(b) == int(binning):
                    return b

        return False

    def _set_camera_param(self, name, size, exp_time, binning, **kwargs):
        camera = self._find_camera(name)

        if self._is_advanced:
            self._tem_csa.Camera = camera

            if not self._tem_csa.Camera.IsInserted():
                self._tem_csa.Camera.Insert()

            settings = self._tem_csa.CameraSettings
            capabilities = settings.Capabilities
            settings.ReadoutArea = size

            binning = self._check_binning(binning, camera, is_advanced=True)
            if binning:
                settings.Binning = binning

            # Set exposure after binning, since it adjusted automatically when binning is set
            settings.ExposureTime = exp_time

            if 'align_image' in kwargs and capabilities.SupportsDriftCorrection:
                settings.AlignImage = kwargs['align_image']
            if 'electron_counting' in kwargs and capabilities.SupportsElectronCounting:
                settings.ElectronCounting = kwargs['electron_counting']
            if 'eer' in kwargs and hasattr(capabilities, 'SupportsEER'):
                self.EER = kwargs['eer']
            if 'frame_ranges' in kwargs:  # a list of tuples
                dfd = settings.DoseFractionsDefinition
                dfd.Clear()
                for i in kwargs['frame_ranges']:
                    dfd.AddRange(i)

            print("Movie of %s frames will be saved to: %s" % (
                settings.CalculateNumberOfFrames(),
                settings.PathToImageStorage + settings.SubPathPattern))

        else:
            info = camera.Info
            settings = camera.AcqParams
            settings.ImageSize = size

            binning = self._check_binning(binning, camera)
            if binning:
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
                    raise Exception("Pre-exposures can only be be done when the shutter mode is set to BOTH")
                settings.PreExposureTime = kwargs['pre_exp_time']
            if 'pre_exp_pause_time' in kwargs:
                if kwargs['shutter_mode'] != AcqShutterMode.BOTH:
                    raise Exception("Pre-exposures can only be be done when the shutter mode is set to BOTH")
                settings.PreExposurePauseTime = kwargs['pre_exp_pause_time']

            # Set exposure after binning, since it adjusted automatically when binning is set
            settings.ExposureTime = exp_time

    def _set_film_param(self, film_text, exp_time, **kwargs):
        self._tem_cam.FilmText = film_text
        self._tem_cam.ManualExposureTime = exp_time

    def _acquire(self, cameraName):
        """ Perform actual acquisition.

        :returns: Image object
        """
        self._tem_acq.RemoveAllAcqDevices()
        self._tem_acq.AddAcqDeviceByName(cameraName)
        img = self._tem_acq.AcquireImages()

        if self._prev_shutter_mode is not None:
            # restore previous shutter mode
            obj = self._prev_shutter_mode[0]
            old_value = self._prev_shutter_mode[1]
            obj.ShutterMode = old_value

        return Image(img[0])

    def _check_prerequisites(self):
        """ Check if buffer cycle or LN filling is running before acquisition call. """
        counter = 0
        while counter < 10:
            if self._tem.Vacuum.PVPRunning:
                print("Buffer cycle in progress, waiting...\r")
                time.sleep(2)
                counter += 1
            else:
                print("Checking buffer levels...")
                break

        counter = 0
        while counter < 40:
            if self._tem.TemperatureControl.DewarsAreBusyFilling:
                print("Dewars are filling, waiting...\r")
                time.sleep(30)
                counter += 1
            else:
                print("Checking dewars levels...")
                break

    def acquire_tem_image(self, cameraName, size, exp_time=1, binning=1, **kwargs):
        """ Acquire a TEM image.

        :param cameraName: Camera name
        :type cameraName: str
        :param size: Image size (AcqImageSize enum)
        :type size: IntEnum
        :param exp_time: Exposure time in seconds
        :type exp_time: float
        :param binning: Binning factor
        :keyword bool align_image: Whether frame alignment (i.e. drift correction) is to be applied to the final image as well as the intermediate images
        :keyword bool electron_counting: Use counting mode
        :keyword bool eer: Use EER mode
        :keyword list frame_ranges: List of frame ranges that define the intermediate images, tuples [(1,2), (2,3)]
        :returns: Image object
        """
        self._set_camera_param(cameraName, size, exp_time, binning, **kwargs)
        if self._is_advanced:
            img = self._tem_csa.Acquire()
            self._check_prerequisites()
            self._tem_csa.Wait()
            return Image(img, isAdvanced=True)

        self._check_prerequisites()
        self._acquire(cameraName)

    def acquire_stem_image(self, cameraName, size, dwell_time=1E-5, binning=1, **kwargs):
        """ Acquire a STEM image.

        :param cameraName: Camera name
        :type cameraName: str
        :param size: Image size (AcqImageSize enum)
        :type size: IntEnum
        :param dwell_time: Dwell time in seconds. The frame time equals the dwell time times the number of pixels plus some overhead (typically 20%, used for the line flyback)
        :type dwell_time: float
        :param binning: Binning factor
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

        settings = self._tem_acq.StemAcqParams
        settings.ImageSize = size

        binning = self._check_binning(binning, det)
        if binning:
            settings.Binning = binning

        settings.DwellTime = dwell_time

        print("Max resolution (?):",
              settings.MaxResolution.X,
              settings.MaxResolution.Y)

        self._check_prerequisites()
        self._acquire(cameraName)

    def acquire_film(self, film_text, exp_time, **kwargs):
        """ Expose a film.

        :param film_text: Film text
        :type film_text: str
        :param exp_time: Exposure time in seconds
        :type exp_time: float
        """
        if self._tem_cam.Stock > 0:
            self._tem_cam.PlateLabelDataType = PlateLabelDateFormat.DDMMYY
            self._tem_cam.ExposureNumber += 1  # TODO: check this
            self._tem_cam.MainScreen = ScreenPosition.UP
            self._tem_cam.ScreenDim = True
            self._set_film_param(film_text, exp_time, **kwargs)
            self._tem_cam.TakeExposure()
        else:
            raise Exception("Plate stock is empty!")


class Detectors:
    """ CCD/DDD, film/plate and STEM detectors. """
    def __init__(self, microscope):
        self._tem_acq = microscope._tem.Acquisition
        self._tem_csa = microscope._tem_adv.Acquisitions.CameraSingleAcquisition
        self._tem_cam = microscope._tem.Camera

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
                "pixel_size(um)": tuple(size / 1e-6 for size in info.PixelSize),
                "binnings": [int(b) for b in info.Binnings],
                "shutter_modes": [AcqShutterMode(x).name for x in info.ShutterModes],
                "pre_exposure_limits(s)": (param.MinPreExposureTime, param.MaxPreExposureTime),
                "pre_exposure_pause_limits(s)": (param.MinPreExposurePauseTime, param.MaxPreExposurePauseTime)
            }
        for cam in self._tem_csa.SupportedCameras:
            self._tem_csa.Camera = cam
            param = self._tem_csa.CameraSettings.Capabilities
            self._cameras[cam.Name] = {
                "type": "CAMERA_ADVANCED",
                "height": cam.Height,
                "width": cam.Width,
                "pixel_size(um)": (cam.PixelSize.Width / 1e-6, cam.PixelSize.Height / 1e-6),
                "binnings": [int(b.Width) for b in param.SupportedBinnings],
                "exposure_time_range(s)": (param.ExposureTimeRange.Begin, param.ExposureTimeRange.End),
                "supports_dose_fractions": param.SupportsDoseFractions,
                "max_number_of_fractions": param.MaximumNumberOfDoseFractions,
                "supports_drift_correction": param.SupportsDriftCorrection,
                "supports_electron_counting": param.SupportsElectronCounting,
                "supports_eer": getattr(param, 'SupportsEER', False)
            }

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
        """ Get fluorescent screen position. """
        return ScreenPosition(self._tem_cam.MainScreen).name

    @screen.setter
    def screen(self, value):
        self._tem_cam.MainScreen = value

    @property
    def film_settings(self):
        """ Returns a dict with film settings. """
        return {
            "stock": self._tem_cam.Stock,
            "exposure_time": self._tem_cam.ManualExposureTime,
            "film_text": self._tem_cam.FilmText,
            "exposure_number": self._tem_cam.ExposureNumber,
            "user_code": self._tem_cam.Usercode,
            "screen_current": self._tem_cam.ScreenCurrent
        }


class Temperature:
    """ LN dewars and temperature controls. """
    def __init__(self, microscope):
        self._tem_temp_control = microscope._tem.TemperatureControl

    def force_refill(self):
        """ Forces LN refill if the level is below 70%, otherwise does nothing. """
        if self._tem_temp_control.TemperatureControlAvailable:
            self._tem_temp_control.ForceRefill()
        else:
            raise Exception("TemperatureControl is not available")

    def dewar_level(self, dewar):
        """ Returns the LN level in a dewar.

        :param dewar: Dewar name (RefrigerantDewar enum)
        :type dewar: IntEnum
        """
        if self._tem_temp_control.TemperatureControlAvailable:
            return self._tem_temp_control.RefrigerantLevel(dewar)
        else:
            raise Exception("TemperatureControl is not available")

    @property
    def is_dewars_filling(self):
        """ Returns TRUE if any of the dewars is currently busy filling. """
        return self._tem_temp_control.DewarsAreBusyFilling

    @property
    def dewars_remaining_time(self):
        """ Returns remaining time until the next dewar refill.
        Returns -1 if no refill is scheduled (e.g. All room temperature, or no
        dewar present).
        """
        return self._tem_temp_control.DewarsRemainingTime


class Autoloader:
    """ Sample loading functions. """
    def __init__(self, microscope):
        self._tem_autoloader = microscope._tem.AutoLoader

    def load_cartridge(self, slot):
        """ Loads the cartridge in the given slot into the microscope. """
        if self._tem_autoloader.AutoLoaderAvailable:
            if self.get_slot_status(slot) != CassetteSlotStatus.OCCUPIED.name:
                raise Exception("Slot %d is not occupied" % slot)
            self._tem_autoloader.LoadCartridge(slot)
        else:
            raise Exception("Autoloader is not available")

    def unload_cartridge(self):
        """ Unloads the cartridge currently in the microscope and puts it back into its
        slot in the cassette.
        """
        if self._tem_autoloader.AutoLoaderAvailable:
            self._tem_autoloader.UnloadCartridge()
        else:
            raise Exception("Autoloader is not available")

    def run_inventory(self):
        """ Performs an inventory of the cassette. """
        if self._tem_autoloader.AutoLoaderAvailable:
            self._tem_autoloader.PerformCassetteInventory()
        else:
            raise Exception("Autoloader is not available")

    def get_slot_status(self, slot):
        """ The status of the slot specified. """
        if self._tem_autoloader.AutoLoaderAvailable:
            status = self._tem_autoloader.SlotStatus(slot)
            return CassetteSlotStatus(status).name
        else:
            raise Exception("Autoloader is not available")

    @property
    def number_of_cassette_slots(self):
        """ The number of slots in a cassette. """
        if self._tem_autoloader.AutoLoaderAvailable:
            return self._tem_autoloader.NumberOfCassetteSlots
        else:
            raise Exception("Autoloader is not available")


class Stage:
    """ Stage functions. """
    def __init__(self, microscope):
        self._tem_stage = microscope._tem.Stage

    def _from_dict(self, position, values):
        axes = 0
        for key, value in values.items():
            if key not in 'xyzab':
                raise ValueError("Unexpected axes: %s" % key)
            attr_name = key.upper()
            setattr(position, attr_name, float(value))
            axes |= getattr(StageAxes, attr_name)
        return axes

    @property
    def status(self):
        """ The current state of the stage. """
        return StageStatus(self._tem_stage.Status).name

    @property
    def holder_type(self):
        """ The current specimen holder type. """
        return StageHolderType(self._tem_stage.Holder).name

    @property
    def position(self):
        """ The current position of the stage. """
        pos = self._tem_stage.Position
        axes = 'xyzab'
        return {key: getattr(pos, key.upper()) for key in axes}

    def go_to(self, speed=None, **kwargs):
        """ Makes the holder directly go to the new position by moving all axes
        simultaneously. Keyword args can be x,y,z,a or b.

        :param speed: fraction of the standard speed setting (max 1.0)
        :type speed: float
        """
        if self._tem_stage.Status == StageStatus.READY:
            pos = self._tem_stage.Position
            axes = self._from_dict(pos, **kwargs)
            if speed:
                self._tem_stage.GoToWithSpeed(axes, speed)
            else:
                self._tem_stage.GoTo(axes)
        else:
            print("Stage is not ready.")

    def move_to(self, **kwargs):
        """ Makes the holder safely move to the new position.
        Keyword args can be x,y,z,a or b.
        """
        if self._tem_stage.Status == StageStatus.READY:
            pos = self._tem_stage.Position
            axes = self._from_dict(pos, **kwargs)
            self._tem_stage.MoveTo(axes)
        else:
            print("Stage is not ready.")

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
            print("PiezoStage interface is not available.")

    @property
    def position(self):
        pos = self._tem_pstage.CurrentPosition
        axes = 'xyz'
        return {key: getattr(pos, key.upper()) for key in axes}

    @property
    def position_range(self):
        return self._tem_pstage.GetPositionRange()

    @property
    def velocity(self):
        pos = self._tem_pstage.CurrentJogVelocity
        axes = 'xyz'
        return {key: getattr(pos, key.upper()) for key in axes}


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
    def is_colvalves_open(self):
        """ The status of the column valves. """
        return self._tem_vacuum.ColumnValvesOpen

    def colvalves_open(self):
        """ Open column valves. """
        self._tem_vacuum.ColumnValvesOpen = True

    def colvalves_close(self):
        """ Close column valves. """
        self._tem_vacuum.ColumnValvesOpen = False

    def run_buffer_cycle(self):
        """ Runs a pumping cycle to empty the buffer. """
        self._tem_vacuum.RunBufferCycle()

    @property
    def gauges(self):
        """ Returns a dict with vacuum gauges information. """
        gauges = {}
        for g in self._tem_vacuum.Gauges:
            g.Read()
            gauges[g.Name] = {
                "status": GaugeStatus(g.Status).name,
                "pressure": g.Pressure,
                "level": GaugePressureLevel(g.PressureLevel).name
            }
        return gauges


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
        """ The current measured on the fluorescent screen (units: Amperes). """
        return self._tem_cam.ScreenCurrent

    @property
    def is_beam_blanked(self):
        """ Status of the beam blanker. """
        return self._tem_illumination.BeamBlanked

    @property
    def is_shutter_override_on(self):
        """ Determines the state of the shutter override function. """
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

    def beam_unblank(self):
        """ Deactivates the beam blanker. """
        self._tem_illumination.BeamBlanked = False

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
    def is_stem_available(self):
        """ Returns whether the microscope has a STEM system or not. """
        return self._tem_control.StemAvailable

    def enable(self):
        """ Switch to STEM mode."""
        self._tem_control.InstrumentMode = InstrumentMode.STEM

    def disable(self):
        """ Switch back to TEM mode. """
        self._tem_control.InstrumentMode = InstrumentMode.TEM

    @property
    def magnification(self):
        """ The magnification value in STEM mode. """
        if self._tem_control.InstrumentMode == InstrumentMode.STEM:
            return self._tem_illumination.StemMagnification

    @magnification.setter
    def magnification(self, mag):
        self._tem_illumination.StemMagnification = mag

    @property
    def rotation(self):
        """ The STEM rotation angle (in radians). """
        if self._tem_control.InstrumentMode == InstrumentMode.STEM:
            return self._tem_illumination.StemRotation

    @rotation.setter
    def rotation(self, rot):
        self._tem_illumination.StemRotation = rot

    @property
    def scan_field_of_view(self):
        """ STEM full scan field of view. """
        if self._tem_control.InstrumentMode == InstrumentMode.STEM:
            return self._tem_illumination.StemFullScanFieldOfView

    @scan_field_of_view.setter
    def scan_field_of_view(self, values):
        vector = self._tem_illumination.StemFullScanFieldOfView
        vector.X = values[0]
        vector.Y = values[1]
        self._tem_illumination.StemFullScanFieldOfView = vector


class Illumination:
    """ Illumination functions. """
    def __init__(self, tem):
        self._tem = tem
        self._tem_illumination = self._tem.Illumination
        self.spotsize = self._tem_illumination.SpotsizeIndex
        self.intensity = self._tem_illumination.Intensity
        self.intensity_zoom = self._tem_illumination.IntensityZoomEnabled
        self.intensity_limit = self._tem_illumination.IntensityLimitEnabled
        self.beam_shift = Vector(self._tem_illumination, 'Shift')
        self.rotation_center = Vector(self._tem_illumination, 'RotationCenter')
        self.condenser_stigmator = Vector(self._tem_illumination, 'CondenserStigmator', range=(-1.0, 1.0))

        if self._tem.Configuration.CondenserLensSystem == CondenserLensSystem.THREE_CONDENSER_LENSES:
            self.illuminated_area = self._tem_illumination.IlluminatedArea
            self.probe_defocus = self._tem_illumination.ProbeDefocus
            self.convergence_angle = self._tem_illumination.ConvergenceAngle
            self.C3ImageDistanceParallelOffset = self._tem_illumination.C3ImageDistanceParallelOffset

    @property
    def mode(self):
        """ Illumination mode: microprobe or nanoprobe. """
        return IlluminationMode(self._tem_illumination.Mode).name

    @mode.setter
    def mode(self, value):
        self._tem_illumination.Mode = value

    @property
    def dark_field_mode(self):
        """ Dark field mode. """
        return DarkFieldMode(self._tem_illumination.DFMode).name

    @dark_field_mode.setter
    def dark_field_mode(self, value):
        self._tem_illumination.DFMode = value

    @property
    def condenser_mode(self):
        """ Mode of the illumination system, parallel or probe. """
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
        Units: radians, either in Cartesian (x,y) or polar (conical)
        tilt angles. The accuracy of the beam tilt physical units
        depends on a calibration of the tilt angles.
        """
        mode = self._tem_illumination.DFMode
        tilt = self._tem_illumination.Tilt
        if mode == DarkFieldMode.CONICAL:
            return tilt[0] * math.cos(tilt[1]), tilt[0] * math.sin(tilt[1])
        elif mode == DarkFieldMode.CARTESIAN:
            return tilt
        else:
            return 0.0, 0.0  # Microscope might return nonsense if DFMode is OFF

    @beam_tilt.setter
    def beam_tilt(self, tilt):
        mode = self._tem_illumination.DFMode
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
        self.focus = self._tem_projection.Focus
        self.magnification_index = self._tem_projection.MagnificationIndex
        self.camera_length_index = self._tem_projection.CameraLengthIndex
        self.image_shift = Vector(self._tem_projection, 'ImageShift')
        self.image_beam_shift = Vector(self._tem_projection, 'ImageBeamShift')  # IS with BS compensation
        self.diffraction_shift = Vector(self._tem_projection, 'DiffractionShift')
        self.diffraction_stigmator = Vector(self._tem_projection, 'DiffractionStigmator', range=(-1.0, 1.0))
        self.objective_stigmator = Vector(self._tem_projection, 'ObjectiveStigmator', range=(-1.0, 1.0))
        self.defocus = self._tem_projection.Defocus
        self.image_beam_tilt = Vector(self._tem_projection, 'ImageBeamTilt')  # BT with diffr sh compensation

    @property
    def magnification(self):
        """ The reference magnification value (screen up setting). """
        return self._tem_projection.Magnification

    @property
    def camera_length(self):
        """ The reference camera length (screen up setting). """
        return self._tem_projection.CameraLength

    @property
    def mode(self):
        """ Main mode of the projection system (either imaging or diffraction). """
        return ProjectionMode(self._tem_projection.Mode).name

    @mode.setter
    def mode(self, mode):
        self._tem_projection.Mode = mode

    @property
    def magnification_range(self):
        """ Submode of the projection system (either LM, M, SA, MH, LAD or D).
        The imaging submode can change when the magnification is changed.
        """
        return ProjectionSubMode(self._tem_projection.SubMode).name

    @property
    def image_rotation(self):
        """ The rotation of the image or diffraction pattern on the
        fluorescent screen with respect to the specimen. Units: radians.
        """
        return self._tem_projection.ImageRotation

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
        self._tem_projection.ResetDefocus()


class Apertures:
    """ Apertures and VPP controls. """
    def __init__(self, microscope):
        self._tem_vpp = microscope._tem_adv.PhasePlate
        try:
            self._tem_apertures = microscope._tem.ApertureMechanismCollection
        except:
            print("Apertures interface is not available. Requires a separate license")

    def _find_aperture(self, name):
        """Find aperture object by name. """
        for ap in self._tem_apertures:
            if MechanismId(ap.Id).name == name:
                return ap
        raise KeyError("No aperture with name %s" % name)

    @property
    def vpp_position(self):
        """ Returns the zero-based index of the current VPP preset position. """
        return self._tem_vpp.GetCurrentPresetPosition

    def vpp_next_position(self):
        """ Goes to the next preset location on the VPP aperture. """
        self._tem_vpp.SelectNextPresetPosition()

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
                    raise Exception("Could not select aperture!")

    @property
    def show_all(self):
        """ Returns a dict with apertures information. """
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
        self.shift = Vector(self._tem_gun, 'Shift', range=(-1.0, 1.0))
        self.tilt = Vector(self._tem_gun, 'Tilt', range=(-1.0, 1.0))
        try:
            self._tem_gun1 = microscope._tem.Gun1
        except:
            print("Gun1 interface is not available. Requires TEM Server 7.10+")
        try:
            self._tem_feg = microscope._tem_adv.Source
        except:
            print("Source/FEG interface is not available.")

    @property
    def voltage_offset(self):
        return self._tem_gun1.HighVoltageOffset

    @voltage_offset.setter
    def voltage_offset(self, offset):
        self._tem_gun1.HighVoltageOffset = offset

    @property
    def feg_state(self):
        return FegState(self._tem_feg.State).name

    @property
    def ht_state(self):
        """ The state of the high tension. (The high tension can be on,
        off or disabled). Disabling/enabling can only be done via
        the button on the system on/off-panel, not via script.
        When switching on the high tension, this function cannot
        check if and when the set high tension value is actually reached.
        """
        return HighTensionState(self._tem_feg.HTState).name

    @ht_state.setter
    def ht_state(self, value):
        self._tem_feg.State = value

    @property
    def voltage(self):
        """ The value of the HT setting as displayed in the TEM user
        interface. Units: kVolts.
        """
        state = self._tem_gun.HTState
        if state == HighTensionState.ON:
            return self._tem_gun.HTValue * 1e-3
        else:
            return 0.0

    @voltage.setter
    def voltage(self, value):
        self._tem_gun.HTValue = value * 1000

    @property
    def voltage_max(self):
        """ The maximum possible value of the HT on this microscope. Units: kVolts. """
        return self._tem_gun.HTMaxValue * 1e-3

    @property
    def voltage_offset_range(self):
        return self._tem_gun1.GetHighVoltageOffsetRange()

    @property
    def beam_current(self):
        return self._tem_feg.BeamCurrent

    @property
    def extractor_voltage(self):
        return self._tem_feg.ExtractorVoltage

    @property
    def focus_index(self):
        focus_index = self._tem_feg.FocusIndex
        return focus_index.Coarse, focus_index.Fine

    def do_flashing(self, flash_type):
        if self._tem_feg.Flashing.IsFlashingAdvised(flash_type):
            self._tem_feg.Flashing.PerformFlashing(flash_type)
        else:
            raise Exception("Flashing type %s is not advised" % flash_type)
