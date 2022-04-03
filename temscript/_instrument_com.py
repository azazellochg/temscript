from .enums import *
from ._properties import *


__all__ = ('GetInstrument', 'Projection', 'CCDCameraInfo', 'CCDAcqParams', 'CCDCamera',
           'STEMDetectorInfo', 'STEMAcqParams', 'STEMDetector', 'AcqImage', 'Acquisition',
           'TemperatureControl', 'AutoLoader', 'UserButton', 'Gauge', 'Vacuum',
           'Stage', 'Camera', 'Illumination', 'Gun', 'BlankerShutter',
           'InstrumentModeControl', 'Configuration', 'Instrument')


class Projection(IUnknown):
    IID = UUID("b39c3ae1-1e41-11d3-ae0e-00a024cba50c")

    Mode = EnumProperty(ProjectionMode, get_index=10, put_index=11)
    Focus = DoubleProperty(get_index=12, put_index=13)
    Magnification = DoubleProperty(get_index=14)
    CameraLength = DoubleProperty(get_index=15)
    MagnificationIndex = LongProperty(get_index=16, put_index=17)
    CameraLengthIndex = LongProperty(get_index=18, put_index=19)
    ImageShift = VectorProperty(get_index=20, put_index=21)
    ImageBeamShift = VectorProperty(get_index=22, put_index=23)
    DiffractionShift = VectorProperty(get_index=24, put_index=25)
    DiffractionStigmator = VectorProperty(get_index=26, put_index=27)
    ObjectiveStigmator = VectorProperty(get_index=28, put_index=29)
    Defocus = DoubleProperty(get_index=30, put_index=31)
    SubModeString = StringProperty(get_index=32)
    SubMode = EnumProperty(ProjectionSubMode, get_index=33)
    SubModeMinIndex = LongProperty(get_index=34)
    SubModeMaxIndex = LongProperty(get_index=35)
    ObjectiveExcitation = DoubleProperty(get_index=36)
    ProjectionIndex = LongProperty(get_index=37, put_index=38)
    LensProgram = EnumProperty(LensProg, get_index=39, put_index=40)
    ImageRotation = DoubleProperty(get_index=41)
    DetectorShift = EnumProperty(ProjectionDetectorShift, get_index=42, put_index=43)
    DetectorShiftMode = EnumProperty(ProjDetectorShiftMode, get_index=44, put_index=45)
    ImageBeamTilt = VectorProperty(get_index=46, put_index=47)

    RESET_DEFOCUS_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT)(7, "ResetDefocus")
    NORMALIZE_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_int)(8, "Normalize")
    CHANGE_PROJECTION_INDEX_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_long)(9, "ChangeProjectionIndex")

    def ResetDefocus(self):
        Projection.RESET_DEFOCUS_METHOD(self.get())

    def Normalize(self, norm):
        Projection.NORMALIZE_METHOD(self.get(), norm)

    def ChangeProjectionIndex(self, add_val):
        Projection.CHANGE_PROJECTION_INDEX_METHOD(self.get(), add_val)


class CCDCameraInfo(IUnknown):
    IID = UUID("024ded60-b124-4514-bfe2-02c0f5c51db9")

    Name = StringProperty(get_index=7)
    Width = LongProperty(get_index=8)
    Height = LongProperty(get_index=9)
    PixelSize = VectorProperty(get_index=10)
    ShutterMode = EnumProperty(AcqShutterMode, get_index=13, put_index=14)
    _ShutterModes = SafeArrayProperty(get_index=12)
    _Binnings = SafeArrayProperty(get_index=11)

    @property
    def Binnings(self):
        return self._Binnings.as_list(int)

    @property
    def ShutterModes(self):
        return self._ShutterModes.as_list(AcqShutterMode)


class CCDAcqParams(IUnknown):
    IID = UUID("c03db779-1345-42ab-9304-95b85789163d")

    ImageSize = EnumProperty(AcqImageSize, get_index=7, put_index=8)
    ExposureTime = DoubleProperty(get_index=9, put_index=10)
    Binning = LongProperty(get_index=11, put_index=12)
    ImageCorrection = EnumProperty(AcqImageCorrection, get_index=13, put_index=14)
    ExposureMode = EnumProperty(AcqExposureMode, get_index=15, put_index=16)
    MinPreExposureTime = DoubleProperty(get_index=17)
    MaxPreExposureTime = DoubleProperty(get_index=18)
    PreExposureTime = DoubleProperty(get_index=19, put_index=20)
    MinPreExposurePauseTime = DoubleProperty(get_index=21)
    MaxPreExposurePauseTime = DoubleProperty(get_index=22)
    PreExposurePauseTime = DoubleProperty(get_index=23, put_index=24)


class CCDCamera(IUnknown):
    IID = UUID("e44e1565-4131-4937-b273-78219e090845")

    Info = ObjectProperty(CCDCameraInfo, get_index=7)
    AcqParams = ObjectProperty(CCDAcqParams, get_index=8, put_index=9)


class STEMDetectorInfo(IUnknown):
    IID = UUID("96de094b-9cdc-4796-8697-e7dd5dc3ec3f")

    Name = StringProperty(get_index=7)
    Brightness = DoubleProperty(get_index=8, put_index=9)
    Contrast = DoubleProperty(get_index=10, put_index=11)
    _Binnings = SafeArrayProperty(get_index=11)

    @property
    def Binnings(self):
        return self._Binnings.as_list(int)


class STEMAcqParams(IUnknown):
    IID = UUID("ddc14710-6152-4963-aea4-c67ba784c6b4")

    ImageSize = EnumProperty(AcqImageSize, get_index=7, put_index=8)
    DwellTime = DoubleProperty(get_index=9, put_index=10)
    Binning = LongProperty(get_index=11, put_index=12)


class STEMDetector(IUnknown):
    __slots__ = '_acquisition'

    IID = UUID("d77c0d65-a1dd-4d0a-af25-c280046a5719")

    Info = ObjectProperty(STEMDetectorInfo, get_index=7)

    def __init__(self, value=None, adopt_reference=False, acquisition=None):
        super(STEMDetector, self).__init__(value=value, adopt_reference=adopt_reference)
        self._acquisition = acquisition

    @property
    def AcqParams(self):
        import warnings
        warnings.warn("The attribute AcqParams of STEMDetector instances is deprecated. Use Acquisition.StemAcqParams instead.", DeprecationWarning)
        return self._acquisition.StemAcqParams

    @AcqParams.setter
    def AcqParams(self, value):
        import warnings
        warnings.warn("The attribute AcqParams of STEMDetector instances is deprecated. Use Acquisition.StemAcqParams instead.", DeprecationWarning)
        self._acquisition.StemAcqParams = value


class AcqImage(IUnknown):
    IID = UUID("e15f4810-43c6-489a-9e8a-588b0949e153")

    Name = StringProperty(get_index=7)
    Width = LongProperty(get_index=8)
    Height = LongProperty(get_index=9)
    Depth = LongProperty(get_index=10)
    _AsSafeArray = SafeArrayProperty(get_index=11)

    AS_FILE_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_wchar_p, ctypes.c_int, ctypes.c_short)(13, "AsFile")

    def AsFile(self, name, format, normalize=False):
        name_bstr = BStr(name)
        format = AcqImageFileFormat[str(format).upper()]
        bool_value = 0xffff if normalize else 0x0000
        AcqImage.AS_FILE_METHOD(self.get(), name_bstr.get(), format, bool_value)

    @property
    def Array(self):
        return self._AsSafeArray.as_array()


class Acquisition(IUnknown):
    IID = UUID("d6bbf89c-22b8-468f-80a1-947ea89269ce")

    _AcquireImages = CollectionProperty(get_index=12, interface=AcqImage)
    Cameras = CollectionProperty(get_index=13, interface=CCDCamera)
    _Detectors = CollectionProperty(get_index=14)

    ADD_ACQ_DEVICE_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(7, "AddAcqDevice")
    ADD_ACQ_DEVICE_BY_NAME_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_wchar_p)(8, "AddAcqDeviceByName")
    REMOVE_ACQ_DEVICE_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(9, "RemoveAcqDevice")
    REMOVE_ACQ_DEVICE_BY_NAME_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_wchar_p)(10, "RemoveAcqDeviceByName")
    REMOVE_ALL_ACQ_DEVICES_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT)(11, "RemoveAllAcqDevices")
    GET_DETECTORS_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(14, "get_Detectors")

    # Methods of STEMDetectors
    GET_ACQ_PARAMS_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(10, "get_AcqParams")
    PUT_ACQ_PARAMS_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(11, "put_AcqParams")

    def AddAcqDevice(self, device):
        if not isinstance(device, (STEMDetector, CCDCamera)):
            raise TypeError("Expected device to be instance of types STEMDetector or CCDCamera")
        Acquisition.ADD_ACQ_DEVICE_METHOD(self.get(), device.get())

    def AddAcqDeviceByName(self, name):
        name_bstr = BStr(name)
        Acquisition.ADD_ACQ_DEVICE_BY_NAME_METHOD(self.get(), name_bstr.get())

    def RemoveAcqDevice(self, device):
        if not isinstance(device, (STEMDetector, CCDCamera)):
            raise TypeError("Expected device to be instance of types STEMDetector or CCDCamera")
        Acquisition.REMOVE_ACQ_DEVICE_METHOD(self.get(), device.get())

    def RemoveAcqDeviceByName(self, name):
        name_bstr = BStr(name)
        Acquisition.REMOVE_ACQ_DEVICE_BY_NAME_METHOD(self.get(), name_bstr.get())

    def RemoveAllAcqDevices(self):
        Acquisition.REMOVE_ALL_ACQ_DEVICES_METHOD(self.get())

    def AcquireImages(self):
        return self._AcquireImages

    @property
    def Detectors(self):
        collection = self._Detectors
        return [STEMDetector(item, acquisition=self) for item in collection]

    @property
    def StemAcqParams(self):
        collection = IUnknown()
        Acquisition.GET_DETECTORS_METHOD(self.get(), collection.byref())
        params = STEMAcqParams()
        Acquisition.GET_ACQ_PARAMS_METHOD(collection.get(), params.byref())
        return params

    @StemAcqParams.setter
    def StemAcqParams(self, value):
        if not isinstance(value, STEMAcqParams):
            raise TypeError("Expected attribute AcqParams to be set to an instance of type STEMAcqParams")
        collection = IUnknown()
        Acquisition.GET_DETECTORS_METHOD(self.get(), collection.byref())
        Acquisition.PUT_ACQ_PARAMS_METHOD(collection.get(), value.get())


class TemperatureControl(IUnknown):
    IID = UUID("71b6e709-b21f-435f-9529-1aee55cfa029")

    TemperatureControlAvailable = VariantBoolProperty(get_index=8)
    DewarsRemainingTime = LongProperty(get_index=10)
    DewarsAreBusyFilling = VariantBoolProperty(get_index=11)

    FORCE_REFILL_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT)(7, "ForceRefill")
    REFRIGERANT_LEVEL_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_int)(9, "RefrigerantLevel")

    def ForceRefill(self):
        TemperatureControl.FORCE_REFILL_METHOD(self.get())

    def RefrigerantLevel(self, rl):
        # rl = RefrigerantLevel
        return TemperatureControl.REFRIGERANT_LEVEL_METHOD(self.get(), rl)


class AutoLoader(IUnknown):
    IID = UUID("28df27ea-2058-41d0-abbd-167fb3bfcd8f")

    AutoLoaderAvailable = VariantBoolProperty(get_index=11)
    NumberOfCassetteSlots = LongProperty(get_index=12)

    LOAD_CARTRIDGE_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_long)(7, "LoadCartridge")
    UNLOAD_CARTRIDGE_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT)(8, "UnloadCartridge")
    PERFORM_CASSETTE_INVENTORY_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT)(9, "PerformCassetteInventory")
    BUFFER_CYCLE_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT)(10, "BufferCycle")
    SLOT_STATUS_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_long)(13, "SlotStatus")

    def LoadCartridge(self, slot):
        AutoLoader.LOAD_CARTRIDGE_METHOD(self.get(), slot)

    def UnloadCartridge(self):
        AutoLoader.UNLOAD_CARTRIDGE_METHOD(self.get())

    def PerformCassetteInventory(self):
        AutoLoader.PERFORM_CASSETTE_INVENTORY_METHOD(self.get())

    def BufferCycle(self):
        AutoLoader.BUFFER_CYCLE_METHOD(self.get())

    def SlotStatus(self, slot):
        # returns CassetteSlotStatus
        return AutoLoader.SLOT_STATUS_METHOD(self.get(), slot)


class Gauge(IUnknown):
    IID = UUID("52020820-18bf-11d3-86e1-00c04fc126dd")

    Name = StringProperty(get_index=8)
    Pressure = DoubleProperty(get_index=9)
    Status = EnumProperty(GaugeStatus, get_index=10)
    PressureLevel = EnumProperty(GaugePressureLevel, get_index=11)

    READ_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT)(7, "Read")

    def Read(self):
        Gauge.READ_METHOD(self.get())


class Vacuum(IUnknown):
    IID = UUID("c7646442-1115-11d3-ae00-00a024cba50c")

    Status = EnumProperty(VacuumStatus, get_index=8)
    PVPRunning = VariantBoolProperty(get_index=9)
    Gauges = CollectionProperty(get_index=10, interface=Gauge)
    ColumnValvesOpen = VariantBoolProperty(get_index=11, put_index=12)

    RUN_BUFFER_CYCLE_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT)(7, "RunBufferCycle")

    def RunBufferCycle(self):
        Vacuum.RUN_BUFFER_CYCLE_METHOD(self.get())


class StagePosition(IUnknown):
    IID = UUID("9851bc4a-1b8c-11d3-ae0a-00a024cba50c")
    AXES = 'xyzab'

    X = DoubleProperty(get_index=9, put_index=10)
    Y = DoubleProperty(get_index=11, put_index=12)
    Z = DoubleProperty(get_index=13, put_index=14)
    A = DoubleProperty(get_index=15, put_index=16)
    B = DoubleProperty(get_index=17, put_index=18)

    def to_dict(self):
        return {key: getattr(self, key.upper()) for key in StagePosition.AXES}

    def from_dict(self, values):
        axes = 0
        for key, value in values.items():
            if key not in StagePosition.AXES:
                raise ValueError("Unexpected axes: %s" % key)
            attr_name = key.upper()
            setattr(self, attr_name, float(value))
            axes |= getattr(StageAxes, attr_name)
        return axes


class StageAxisData(IUnknown):
    IID = UUID("8f1e91c2-b97d-45b8-87c9-423f5eb10b8a")

    MinPos = DoubleProperty(get_index=7)
    MaxPos = DoubleProperty(get_index=8)
    UnitType = EnumProperty(MeasurementUnitType, get_index=9)


class Stage(IUnknown):
    IID = UUID("e7ae1e41-1bf8-11d3-ae0b-00a024cba50c")

    STAGEAXES_FROM_AXIS = {
        'x': StageAxes.X,
        'y': StageAxes.Y,
        'z': StageAxes.Z,
        'a': StageAxes.A,
        'b': StageAxes.B,
    }

    UNIT_STRING = {
        MeasurementUnitType.METERS: 'meters',
        MeasurementUnitType.RADIANS: 'radians'
    }

    Status = EnumProperty(StageStatus, get_index=9)
    _Position = ObjectProperty(get_index=10, interface=StagePosition)
    Holder = EnumProperty(StageHolderType, get_index=11)

    GOTO_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p, ctypes.c_int)(7, 'GoTo')
    MOVETO_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p, ctypes.c_int)(8, 'MoveTo')
    AXIS_DATA_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_int, ctypes.c_void_p)(12, 'AxisData')
    GOTO_WITH_SPEED_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p, ctypes.c_int, ctypes.c_double)(13, 'GoToWithSpeed')

    @property
    def Position(self):
        return self._Position.to_dict()

    def GoTo(self, speed=None, **kw):
        pos = self._Position
        axes = pos.from_dict(kw)
        if not axes:
            return
        if speed is not None:
            Stage.GOTO_WITH_SPEED_METHOD(self.get(), pos.get(), axes, speed)
        else:
            Stage.GOTO_METHOD(self.get(), pos.get(), axes)
            
    def MoveTo(self, **kw):
        pos = self._Position
        axes = pos.from_dict(kw)
        if not axes:
            return
        Stage.MOVETO_METHOD(self.get(), pos.get(), axes)

    def AxisData(self, axis):
        try:
            mask = Stage.STAGEAXES_FROM_AXIS[axis]
        except KeyError:
            raise ValueError("Expected axis name: 'x', 'y', 'z', 'a', or 'b'")
        data = StageAxisData()
        Stage.AXIS_DATA_METHOD(self.get(), mask, data.byref())
        return (data.MinPos, data.MaxPos, Stage.UNIT_STRING.get(data.UnitType))


class Camera(IUnknown):
    IID = UUID("9851bc41-1b8c-11d3-ae0a-00a024cba50c")

    Stock = LongProperty(get_index=8)
    MainScreen = EnumProperty(ScreenPosition, get_index=9, put_index=10)
    IsSmallScreenDown = VariantBoolProperty(get_index=11)
    MeasuredExposureTime = DoubleProperty(get_index=12)
    FilmText = StringProperty(get_index=13, put_index=14)
    ManualExposureTime = DoubleProperty(get_index=15, put_index=16)
    PlateuMarker = VariantBoolProperty(get_index=17, put_index=18)
    ExposureNumber = LongProperty(get_index=19, put_index=20)
    Usercode = StringProperty(get_index=21, put_index=22)
    ManualExposure = VariantBoolProperty(get_index=23, put_index=24)
    PlateLabelDataType = EnumProperty(PlateLabelDateFormat, get_index=25, put_index=26)
    ScreenDim = VariantBoolProperty(get_index=27, put_index=28)
    ScreenDimText = StringProperty(get_index=29, put_index=30)
    ScreenCurrent = DoubleProperty(get_index=31)

    TAKE_EXPOSURE_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT)(7, "TakeExposure")

    def TakeExposure(self):
        Camera.TAKE_EXPOSURE_METHOD(self.get())


class Illumination(IUnknown):
    IID = UUID("ef960690-1c38-11d3-ae0b-00a024cba50c")

    Mode = EnumProperty(IlluminationMode, get_index=8, put_index=9)
    SpotsizeIndex = LongProperty(get_index=10, put_index=11)
    Intensity = DoubleProperty(get_index=12, put_index=13)
    IntensityZoomEnabled = VariantBoolProperty(get_index=14, put_index=15)
    IntensityLimitEnabled = VariantBoolProperty(get_index=16, put_index=17)
    BeamBlanked = VariantBoolProperty(get_index=18, put_index=19)
    Shift = VectorProperty(get_index=20, put_index=21)
    Tilt = VectorProperty(get_index=22, put_index=23)
    RotationCenter = VectorProperty(get_index=24, put_index=25)
    CondenserStigmator = VectorProperty(get_index=26, put_index=27)
    DFMode = EnumProperty(DarkFieldMode, get_index=28, put_index=29)
    #DarkFieldMode = EnumProperty(DarkFieldMode, get_index=28, put_index=29)
    CondenserMode = EnumProperty(CondenserMode, get_index=30, put_index=31)
    IlluminatedArea = DoubleProperty(get_index=32, put_index=33)
    ProbeDefocus = DoubleProperty(get_index=34)
    ConvergenceAngle = DoubleProperty(get_index=35)
    StemMagnification = DoubleProperty(get_index=36, put_index=37)
    StemRotation = DoubleProperty(get_index=38, put_index=39)
    StemFullScanFieldOfView = VectorProperty(get_index=40, put_index=41)
    C3ImageDistanceParallelOffset = DoubleProperty(get_index=42, put_index=43)

    NORMALIZE_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_int)(7, "Normalize")

    def Normalize(self, norm):
        Illumination.NORMALIZE_METHOD(self.get(), norm)


class Gun(IUnknown):
    IID = UUID("e6f00870-3164-11d3-b4c8-00a024cb9221")

    HTState = EnumProperty(HighTensionState, get_index=7, put_index=8)
    HTValue = DoubleProperty(get_index=9, put_index=10)
    HTMaxValue = DoubleProperty(get_index=11)
    Shift = VectorProperty(get_index=12, put_index=13)
    Tilt = VectorProperty(get_index=14, put_index=15)


class Gun1(IUnknown):
    IID = UUID("?")

    HighVoltageOffset = DoubleProperty(get_index=8, put_index=9)

    GET_HIGH_VOLTAGE_OFFSET_RANGE_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT)(7, "GetHighVoltageOffsetRange")

    def GetHighVoltageOffsetRange(self):
        return Gun1.GET_HIGH_VOLTAGE_OFFSET_RANGE_METHOD(self.get())


class BlankerShutter(IUnknown):
    IID = UUID("f1f59bb0-f8a0-439d-a3bf-87f527b600c4")

    ShutterOverrideOn = VariantBoolProperty(get_index=7, put_index=8)


class UserButton(IUnknown):
    IID = UUID("e6f00871-3164-11d3-b4c8-00a024cb9221")

    Name = StringProperty(get_index=7)
    Label = StringProperty(get_index=8)
    Assignment = StringProperty(get_index=9, put_index=10)


class InstrumentModeControl(IUnknown):
    IID = UUID("8dc0fc71-ff15-40d8-8174-092218d8b76b")

    StemAvailable = VariantBoolProperty(get_index=7)
    InstrumentMode = EnumProperty(InstrumentMode, get_index=8, put_index=9)


class Configuration(IUnknown):
    IID = UUID("39cacdaf-f47c-4bbf-9ffa-a7a737664ced")

    ProductFamily = EnumProperty(ProductFamily, get_index=7)
    CondenserLensSystem = EnumProperty(CondenserLensSystem, get_index=8)


class Aperture(IUnknown):
    IID = UUID("cbf4e5b8-378d-43dd-9c58-f588d5e3444b")

    Name = StringProperty(get_index=7)
    Type = EnumProperty(ApertureType, get_index=8)
    Diameter = DoubleProperty(get_index=9)


class ApertureMechanism(IUnknown):
    IID = UUID("86c13ce3-934e-47de-a211-2009e10e1ee1")

    ApertureCollection = CollectionProperty(get_index=7, interface=Aperture)
    Id = EnumProperty(MechanismId, get_index=8)
    SelectedAperture = ObjectProperty(Aperture, get_index=10)
    State = EnumProperty(MechanismState, get_index=11)
    IsRetractable = VariantBoolProperty(get_index=12)

    SELECT_APERTURE_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(9, "SelectAperture")
    RETRACT_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT)(13, "Retract")
    ENABLE_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT)(14, "Enable")
    DISABLE_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT)(15, "Disable")

    def SelectAperture(self, aperture):
        if not isinstance(aperture, Aperture):
            raise TypeError("Expected aperture to be instance of types Aperture")
        ApertureMechanism.SELECT_APERTURE_METHOD(self.get(), aperture.get())

    def RetractAperture(self):
        ApertureMechanism.RETRACT_METHOD(self.get())

    def Enable(self):
        ApertureMechanism.ENABLE_METHOD(self.get())

    def Disable(self):
        ApertureMechanism.DISABLE_METHOD(self.get())


class Instrument(IUnknown):
    IID = UUID("bc0a2b11-10ff-11d3-ae00-00a024cba50c")

    AutoNormalizeEnabled = VariantBoolProperty(get_index=8, put_index=9)
    Vacuum = ObjectProperty(Vacuum, get_index=13)
    Camera = ObjectProperty(Camera, get_index=14)
    Stage = ObjectProperty(Stage, get_index=15)
    Illumination = ObjectProperty(Illumination, get_index=16)
    Projection = ObjectProperty(Projection, get_index=17)
    Gun = ObjectProperty(Gun, get_index=18)
    UserButtons = CollectionProperty(get_index=19, interface=UserButton)
    AutoLoader = ObjectProperty(AutoLoader, get_index=20)
    TemperatureControl = ObjectProperty(TemperatureControl, get_index=21)
    BlankerShutter = ObjectProperty(BlankerShutter, get_index=22)
    InstrumentModeControl = ObjectProperty(InstrumentModeControl, get_index=23)
    Acquisition = ObjectProperty(Acquisition, get_index=24)
    Configuration = ObjectProperty(Configuration, get_index=25)
    ApertureMechanismCollection = EnumProperty(ApertureMechanism, get_index=26)
    #Gun1 = ObjectProperty(Gun1, get_index=27, put_index=28)

    NORMALIZE_ALL_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT)(7, "NormalizeAll")

    def NormalizeAll(self):
        Instrument.NORMALIZE_ALL_METHOD(self.get())


# TEMScripting.Instrument.1
CLSID_INSTRUMENT = UUID("02CDC9A1-1F1D-11D3-AE11-00A024CBA50C")


def GetInstrument():
    """Returns Instrument instance."""
    instrument = co_create_instance(CLSID_INSTRUMENT, CLSCTX_ALL, Instrument)
    return instrument
