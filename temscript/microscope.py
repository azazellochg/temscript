from .base_microscope import BaseMicroscope, Image, Vector
from .utils.enums import *


class Microscope(BaseMicroscope):
    """ Main class that connects to the COM interface. """

    def __init__(self, address=None, timeout=None, simulate=False):
        super().__init__(address, timeout, simulate)

        self.acquisition = Acquisition(self)
        self.detectors = Detectors(self)
        self.gun = Gun(self)
        self.optics = Optics(self)
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
        self._tem_door = microscope._tem_adv.UserDoorHatch

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
        self._tem_acq = microscope._tem.Acquisition
        self._tem_csa = microscope._tem_adv.Acquisitions.CameraSingleAcquisition
        self._tem_cam = microscope._tem.Camera
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
        raise KeyError(f"No camera with name {name}")

    def _find_stem_detector(self, name):
        """Find STEM detector object by name"""
        for stem in self._tem_acq.Detectors:
            if stem.Info.Name == name:
                return stem
        raise KeyError(f"No STEM detector with name {name}")

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

            print(f"Movie of {settings.CalculateNumberOfFrames()} frames will be "
                  f"saved to: {settings.PathToImageStorage + settings.SubPathPattern}")

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
                settings.PreExposureTime = kwargs['pre_exp_time']
            if 'pre_exp_pause_time' in kwargs:
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

    def acquire_tem_image(self, cameraName, size, exp_time=1, binning=1, **kwargs):
        """ Acquire a TEM image.

        :param cameraName: Camera name
        :type cameraName: str
        :param size: Image size (AcqImageSize enum)
        :type size: IntEnum
        :param exp_time: Exposure time in seconds
        :type exp_time: float
        :param binning: Binning factor
        :returns: Image object
        """
        self._set_camera_param(cameraName, size, exp_time, binning, **kwargs)
        if self._is_advanced:
            img = self._tem_csa.Acquire()
            self._tem_csa.Wait()
            return Image(img, isAdvanced=True)

        self._acquire(cameraName)

    def acquire_stem_image(self, cameraName, size, dwell_time=1E-6, binning=1, **kwargs):
        """ Acquire a STEM image.

        :param cameraName: Camera name
        :type cameraName: str
        :param size: Image size (AcqImageSize enum)
        :type size: IntEnum
        :param dwell_time: Dwell time in seconds. The frame time equals the dwell time times the number of pixels plus some overhead (typically 20%, used for the line flyback)
        :type dwell_time: float
        :param binning: Binning factor
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

    def cameras(self):
        """ Returns a dict with camera parameters. """
        self.cameras = dict()
        for cam in self._tem_acq.Cameras:
            info = cam.Info
            param = cam.AcqParams
            name = info.Name
            self.cameras[name] = {
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
            self.cameras[cam.Name] = {
                "type": "CAMERA_ADVANCED",
                "height": cam.Height,
                "width": cam.Width,
                "pixel_size(um)": (cam.PixelSize.Width / 1e-6, cam.PixelSize.Height / 1e-6),
                "binnings": [int(b.Width) for b in param.SupportedBinnings],
                "exposure_time_range(s)": (param.ExposureTimeRange.Begin, param.ExposureTimeRange.End),
                "supports_dose_fractions": param.SupportsDoseFractions,
                "supports_drift_correction": param.SupportsDriftCorrection,
                "supports_electron_counting": param.SupportsElectronCounting,
                "supports_eer": getattr(param, 'SupportsEER', False)
            }

        return self.cameras

    def stem_detectors(self):
        """ Returns a dict with STEM detectors parameters. """
        self.stem_detectors = dict()
        for det in self._tem_acq.Detectors:
            info = det.Info
            name = info.Name
            self.stem_detectors[name] = {
                "type": "STEM_DETECTOR",
                "binnings": [int(b) for b in info.Binnings],
                "brightness": info.Brightness,
                "contrast": info.Contrast
            }
        return self.stem_detectors

    @property
    def screen(self):
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
        if self._tem_temp_control.TemperatureControlAvailable:
            self._tem_temp_control.ForceRefill()
        else:
            raise Exception("TemperatureControl is not available")

    def dewar_level(self, dewar):
        """ Returns LN level in the dewar.

        :param dewar: Dewar name (RefrigerantDewar enum)
        :type dewar: IntEnum
        """
        if self._tem_temp_control.TemperatureControlAvailable:
            return self._tem_temp_control.RefrigerantLevel(dewar)
        else:
            raise Exception("TemperatureControl is not available")

    @property
    def is_dewars_filling(self):
        return self._tem_temp_control.DewarsAreBusyFilling

    @property
    def dewars_remaining_time(self):
        return self._tem_temp_control.DewarsRemainingTime


class Autoloader:
    """ Sample loading functions. """

    def __init__(self, microscope):
        self._tem_autoloader = microscope._tem.AutoLoader

    def load_cartridge(self, slot):
        if self._tem_autoloader.AutoLoaderAvailable:
            self._tem_autoloader.LoadCartridge(slot)
        else:
            raise Exception("Autoloader is not available")

    def unload_cartridge(self):
        if self._tem_autoloader.AutoLoaderAvailable:
            self._tem_autoloader.UnloadCartridge()
        else:
            raise Exception("Autoloader is not available")

    def run_inventory(self):
        if self._tem_autoloader.AutoLoaderAvailable:
            self._tem_autoloader.PerformCassetteInventory()
        else:
            raise Exception("Autoloader is not available")

    def get_slot_status(self, slot):
        if self._tem_autoloader.AutoLoaderAvailable:
            status = self._tem_autoloader.SlotStatus(slot)
            return CassetteSlotStatus(status).name
        else:
            raise Exception("Autoloader is not available")

    @property
    def number_of_cassette_slots(self):
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
        return StageStatus(self._tem_stage.Status).name

    @property
    def holder_type(self):
        return StageHolderType(self._tem_stage.Holder).name

    @property
    def position(self):
        pos = self._tem_stage.Position
        axes = 'xyzab'
        return {key: getattr(pos, key.upper()) for key in axes}

    def go_to(self, speed=None, **kwargs):
        pos = self._tem_stage.Position
        axes = self._from_dict(pos, kwargs)
        if speed:
            self._tem_stage.GoToWithSpeed(axes, speed)
        else:
            self._tem_stage.GoTo(axes)

    def move_to(self, **kwargs):
        pos = self._tem_stage.Position
        axes = self._from_dict(pos, kwargs)
        self._tem_stage.MoveTo(axes)

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
        self._tem_pstage = microscope._tem_adv.PiezoStage
        self.high_resolution = self._tem_pstage.HighResolution

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
        return VacuumStatus(self._tem_vacuum.Status).name

    @property
    def is_buffer_running(self):
        return self._tem_vacuum.PVPRunning

    @property
    def is_column_open(self):
        return self._tem_vacuum.ColumnValvesOpen

    def run_buffer_cycle(self):
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
        self._tem_illumination = self._tem.Illumination
        self._tem_projection = self._tem.Projection
        self._tem_control = self._tem.InstrumentModeControl
        self.illumination = Illumination(self._tem_illumination)
        self.projection = Projection(self._tem_projection)

    @property
    def is_beam_blanked(self):
        return self._tem_illumination.BeamBlanked

    @property
    def is_shutter_override_on(self):
        return self._tem.BlankerShutter.ShutterOverrideOn

    @property
    def mode(self):
        return InstrumentMode(self._tem_control.InstrumentMode).name

    @mode.setter
    def mode(self, value):
        self._tem_control.InstrumentMode = value

    @property
    def is_stem_available(self):
        return self._tem_control.StemAvailable

    @property
    def is_shutter_override(self):
        return self._tem.BlankerShutter

    @property
    def is_autonormalize_on(self):
        return self._tem.AutoNormalizeEnabled

    def normalize_all(self):
        """ Normalize all lenses. """
        self._tem.NormalizeAll()

    def normalize(self, mode):
        if mode in ProjectionNormalization:
            self._tem_projection.Normalize(mode)
        elif mode in IlluminationNormalization:
            self._tem_illumination.Normalize(mode)
        else:
            raise ValueError(f"Unknown normalization mode: {mode}")


class Illumination:
    """ Illumination functions. """

    def __init__(self, illumination):
        self._tem_illumination = illumination
        self.spotsize = self._tem_illumination.SpotsizeIndex
        self.intensity = self._tem_illumination.Intensity
        self.intensity_zoom = self._tem_illumination.IntensityZoomEnabled
        self.intensity_limit = self._tem_illumination.IntensityLimitEnabled
        self.shift = Vector(self._tem_illumination, 'Shift')
        self.tilt = Vector(self._tem_illumination, 'Tilt')
        self.rotation_center = Vector(self._tem_illumination, 'RotationCenter')
        self.condenser_stigmator = Vector(self._tem_illumination, 'CondenserStigmator', range=(-1.0, 1.0))
        self.illuminated_area = self._tem_illumination.IlluminatedArea
        self.probe_defocus = self._tem_illumination.ProbeDefocus
        self.convergence_angle = self._tem_illumination.ConvergenceAngle
        self.stem_magnification = self._tem_illumination.StemMagnification
        self.stem_rotation = self._tem_illumination.StemRotation
        self.stem_fov = Vector(self._tem_illumination, 'StemFullScanFieldOfView')
        self.C3_ImageDistance_ParallelOffset = self._tem_illumination.C3ImageDistanceParallelOffset

    @property
    def mode(self):
        return IlluminationMode(self._tem_illumination.Mode).name

    @mode.setter
    def mode(self, value):
        self._tem_illumination.Mode = value

    @property
    def dark_field_mode(self):
        return DarkFieldMode(self._tem_illumination.DFMode).name

    @dark_field_mode.setter
    def dark_field_mode(self, value):
        self._tem_illumination.DFMode = value

    @property
    def condenser_mode(self):
        return CondenserMode(self._tem_illumination.CondenserMode).name

    @condenser_mode.setter
    def condenser_mode(self, value):
        self._tem_illumination.CondenserMode = value


class Projection:
    """ Projection system functions. """

    def __init__(self, projection):
        self._tem_projection = projection
        self.focus = self._tem_projection.Focus
        self.magnification_index = self._tem_projection.MagnificationIndex
        self.camera_length_index = self._tem_projection.CameraLengthIndex
        self.image_shift = Vector(self._tem_projection, 'ImageShift')
        self.image_beam_shift = Vector(self._tem_projection, 'ImageBeamShift')
        self.diffraction_shift = Vector(self._tem_projection, 'DiffractionShift')
        self.diffraction_stigmator = Vector(self._tem_projection, 'DiffractionStigmator', range=(-1.0, 1.0))
        self.objective_stigmator = Vector(self._tem_projection, 'ObjectiveStigmator', range=(-1.0, 1.0))
        self.defocus = self._tem_projection.Defocus
        self.image_beam_tilt = Vector(self._tem_projection, 'ImageBeamTilt')

    @property
    def magnification(self):
        return self._tem_projection.Magnification

    @property
    def camera_length(self):
        return self._tem_projection.CameraLength

    @property
    def magnification_range(self):
        return ProjectionSubMode(self._tem_projection.SubMode).name

    @property
    def image_rotation(self):
        return self._tem_projection.ImageRotation

    @property
    def is_eftem_on(self):
        return LensProg(self._tem_projection.LensProgram) == LensProg.EFTEM

    def eftem_on(self):
        self._tem_projection.LensProgram = LensProg.EFTEM

    def reset_defocus(self):
        self._tem_projection.ResetDefocus()


class Apertures:
    """ Apertures and VPP controls. """

    def __init__(self, microscope):
        self._tem_apertures = microscope._tem.ApertureMechanismCollection
        self._tem_vpp = microscope._tem_adv.PhasePlate

    def _find_aperture(self, name):
        """Find aperture object by name. """
        for ap in self._tem_apertures:
            if MechanismId(ap.Id).name == name:
                return ap
        raise KeyError(f"No aperture with name {name}")

    @property
    def vpp_position(self):
        return self._tem_vpp.GetCurrentPresetPosition

    def vpp_next_position(self):
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

        :param aperture: Aperture name
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
        self._has_feg = self._init_feg(microscope)
        self._has_gun1 = self._init_gun1(microscope)
        self.shift = Vector(self._tem_gun, 'Shift', range=(-1.0, 1.0))
        self.tilt = Vector(self._tem_gun, 'Tilt', range=(-1.0, 1.0))

    def _init_gun1(self, microscope):
        try:
            self._tem_gun1 = microscope._tem.Gun1
            self.voltage_offset = self._tem_gun1.HighVoltageOffset
            return True
        except:
            print("Gun1 interface is not available. Requires TEM Server 7.10+")
            return False

    def _init_feg(self, microscope):
        try:
            self._tem_feg = microscope._tem_adv.Source
            return True
        except:
            print("Source/FEG interface is not available.")
            return False

    @property
    def feg_state(self):
        if self._has_feg:
            return FegState(self._tem_feg.State).name

    @property
    def ht_state(self):
        return HighTensionState(self._tem_feg.HTState).name

    @ht_state.setter
    def ht_state(self, value):
        self._tem_feg.State = value

    @property
    def voltage(self):
        state = self._tem_gun.HTState
        if state == HighTensionState.ON:
            return self._tem_gun.HTValue * 1e-3
        else:
            return 0.0

    @voltage.setter
    def voltage(self, value):
        self._tem_gun.HTValue = value

    @property
    def voltage_max(self):
        return self._tem_gun.HTMaxValue

    @property
    def voltage_offset_range(self):
        if self._has_gun1:
            return self._tem_gun1.GetHighVoltageOffsetRange()

    @property
    def beam_current(self):
        if self._has_feg:
            return self._tem_feg.BeamCurrent

    @property
    def extractor_voltage(self):
        if self._has_feg:
            return self._tem_feg.ExtractorVoltage

    @property
    def focus_index(self):
        if self._has_feg:
            focus_index = self._tem_feg.FocusIndex
            return focus_index.Coarse, focus_index.Fine

    def do_flashing(self, flash_type):
        if self._has_feg:
            if self._tem_feg.Flashing.IsFlashingAdvised(flash_type):
                self._tem_feg.Flashing.PerformFlashing(flash_type)
            else:
                raise Exception(f"Flashing type {flash_type} is not advised")
