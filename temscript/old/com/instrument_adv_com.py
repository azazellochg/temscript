from temscript.utils.enums import *
from temscript.old.properties import *


__all__ = ('GetAdvancedInstrument', 'CameraAcquisitionCapabilities', 'CameraSettings',
           'AdvancedCamera', 'AcquiredImage', 'CameraSingleAcquisition',
           'Acquisitions', 'Phaseplate', 'PiezoStagePosition',
           'PiezoStageVelocity', 'PiezoStage', 'UserDoorHatch',
           'Feg', 'AdvancedInstrument', 'CameraSettingsInternal',
           'AdvancedInstrumentInternal')


class Binning(IUnknown):
    IID = UUID("3e34a756-bd44-405f-85a8-83496341476c")

    Width = LongProperty(get_index=7)
    Height = LongProperty(get_index=8)


class TimeRange(IUnknown):
    IID = UUID("07347e4b-ddb5-4971-a1e4-999bee10de08")

    Begin = DoubleProperty(get_index=7)
    End = DoubleProperty(get_index=8)


class FrameRange(IUnknown):
    IID = UUID("85312757-930d-4aa7-8f5c-627f4c65c3bf")

    Begin = LongProperty(get_index=7, put_index=8)
    End = LongProperty(get_index=9, put_index=10)


class FrameRangeList(IUnknown):
    IID = UUID("91aa6644-a9c7-4905-812b-6008b79fe966")

    ADD_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(10, "Add")
    ADD_RANGE_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_long, ctypes.c_long)(11, "AddRange")
    CLEAR_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT)(12, "Clear")

    def Add(self, value):
        if isinstance(value, tuple) or isinstance(value, list):
            rng = FrameRange()
            rng.Begin, rng.End = value[0], value[1]
            FrameRangeList.ADD_METHOD(self.get(), rng.get())
        else:
            raise TypeError("Input value must be a tuple or a list!")

    def AddRange(self, begin, end):
        FrameRangeList.ADD_RANGE_METHOD(self.get(), int(begin), int(end))

    def Clear(self):
        FrameRangeList.CLEAR_METHOD(self.get())


class KeyValuePair(IUnknown):
    IID = UUID("565dfad5-a223-44c1-bc09-22c450b21d24")

    Key = StringProperty(get_index=7)
    ValueAsString = StringProperty(get_index=8)


class PixelSize(IUnknown):
    IID = UUID("38d86b8f-8681-4eaf-b4d2-a7312a0fb981")

    Width = DoubleProperty(get_index=7, put_index=8)
    Height = DoubleProperty(get_index=9, put_index=10)


class CameraAcquisitionCapabilities(IUnknown):
    IID = UUID("c4c83905-0d53-47f3-8ae4-91f1248aa0f9")

    SupportedBinnings = NewCollectionProperty(get_index=7, interface=Binning)
    ExposureTimeRange = ObjectProperty(TimeRange, get_index=8)
    SupportsDoseFractions = VariantBoolProperty(get_index=9)
    MaximumNumberOfDoseFractions = LongProperty(get_index=10)
    SupportsDriftCorrection = VariantBoolProperty(get_index=11)
    SupportsElectronCounting = VariantBoolProperty(get_index=12)
    _SupportsEER = VariantBoolProperty(get_index=13)

    @property
    def SupportsEER(self):
        """ Older advanced tem scripting does not have this property. """
        try:
            eer = self._SupportsEER
            return eer
        except:
            return False


class CameraSettings(IUnknown):
    IID = UUID("f4e2613d-2e3a-4be4-b9e1-4bb1831d3eb1")

    Capabilities = ObjectProperty(CameraAcquisitionCapabilities, get_index=7)
    PathToImageStorage = StringProperty(get_index=8)
    SubPathPattern = StringProperty(get_index=9, put_index=10)
    ExposureTime = DoubleProperty(get_index=11, put_index=12)
    ReadoutArea = EnumProperty(AcqImageSize, get_index=13, put_index=14)
    Binning = ObjectProperty(Binning, get_index=15, put_index=16)
    DoseFractionsDefinition = ObjectProperty(FrameRangeList, get_index=17)
    AlignImage = VariantBoolProperty(get_index=18, put_index=19)
    ElectronCounting = VariantBoolProperty(get_index=20, put_index=21)
    _EER = VariantBoolProperty(get_index=23, put_index=24)

    CALCULATE_NUMBER_OF_FRAMES_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(22, "CalculateNumberOfFrames")

    def CalculateNumberOfFrames(self):
        res = ctypes.c_long(-1)
        CameraSettings.CALCULATE_NUMBER_OF_FRAMES_METHOD(self.get(), ctypes.byref(res))
        return res.value

    @property
    def EER(self):
        """ Older advanced tem scripting does not have this property. """
        try:
            eer = self._EER
            return eer
        except:
            return False


class AdvancedCamera(IUnknown):
    IID = UUID("ba60831e-7a06-4f1d-abe9-80d26227fcb9")

    Name = StringProperty(get_index=7)
    Width = LongProperty(get_index=8)
    Height = LongProperty(get_index=9)
    PixelSize = ObjectProperty(PixelSize, get_index=10)
    IsInserted = VariantBoolProperty(get_index=11)

    INSERT_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT)(12, "Insert")
    RETRACT_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT)(13, "Retract")

    def Insert(self):
        AdvancedCamera.INSERT_METHOD(self.get())

    def Retract(self):
        AdvancedCamera.RETRACT_METHOD(self.get())


class AcquiredImage(IUnknown):
    IID = UUID("88febcf4-397b-4723-aeba-3cacc4ef6840")

    Width = LongProperty(get_index=7)
    Height = LongProperty(get_index=8)
    PixelType = EnumProperty(ImagePixelType, get_index=9)
    BitDepth = LongProperty(get_index=10)
    _Metadata = NewCollectionProperty(get_index=11)
    _AsSafeArray = SafeArrayProperty(get_index=12)

    SAVE_TO_FILE_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_wchar_p, ctypes.c_short)(14, "SaveToFile")

    def SaveToFile(self, filePath, normalize=False):
        name_bstr = BStr(filePath)
        bool_value = 0xffff if normalize else 0x0000
        AcquiredImage.SAVE_TO_FILE_METHOD(self.get(), name_bstr.get(), bool_value)

    @property
    def Metadata(self):
        collection = self._Metadata
        md_list = [KeyValuePair(item) for item in collection]
        return {i.Key: i.ValueAsString for i in md_list}

    @property
    def Array(self):
        return self._AsSafeArray.as_array()


class CameraSingleAcquisition(IUnknown):
    IID = UUID("a927ea10-74b0-45a9-8368-fcdd52498053")

    SupportedCameras = NewCollectionProperty(get_index=7, interface=AdvancedCamera)
    Camera = ObjectProperty(AdvancedCamera, get_index=None, put_index=8)
    CameraSettings = ObjectProperty(CameraSettings, get_index=9)
    IsActive = VariantBoolProperty(get_index=11)

    ACQUIRE_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(10, "Acquire")
    WAIT_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT)(12, "Wait")

    @property
    def Acquire(self):
        image = AcquiredImage()
        CameraSingleAcquisition.ACQUIRE_METHOD(self.get(), image.byref())
        return image

    def Wait(self):
        CameraSingleAcquisition.WAIT_METHOD(self.get())


class Acquisitions(IUnknown):
    IID = UUID("27f7ddc7-bad9-4e2e-b9b6-e7644eb152ec")

    Cameras = NewCollectionProperty(get_index=7, interface=AdvancedCamera)
    CameraSingleAcquisition = ObjectProperty(CameraSingleAcquisition, get_index=8)


class Phaseplate(IUnknown):
    IID = UUID("2605f3d9-9365-42fb-8b09-78f8d6b114d4")

    GetCurrentPresetPosition = LongProperty(get_index=8)

    SELECT_NEXT_PRESET_POSITION_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT)(7, "SelectNextPresetPosition")

    def SelectNextPresetPosition(self):
        Phaseplate.SELECT_NEXT_PRESET_POSITION_METHOD(self.get())


class PiezoStagePosition(IUnknown):
    IID = UUID("7609fe63-eb1c-4e03-8f41-fa17dbd26cae")

    X = DoubleProperty(get_index=7, put_index=8)
    Y = DoubleProperty(get_index=9, put_index=10)
    Z = DoubleProperty(get_index=11, put_index=12)


class PiezoStageVelocity(IUnknown):
    IID = UUID("7d5b599a-000f-4d37-93b8-708896109200")

    X = DoubleProperty(get_index=7, put_index=8)
    Y = DoubleProperty(get_index=9, put_index=10)
    Z = DoubleProperty(get_index=11, put_index=12)


class PiezoStage(IUnknown):
    IID = UUID("d7d3091c-2b98-495b-af47-78990e19a0a9")

    CurrentPosition = ObjectProperty(PiezoStagePosition, get_index=12)
    CurrentJogVelocity = ObjectProperty(PiezoStageVelocity, get_index=13)
    AvailableAxes = LongProperty(get_index=15)
    HighResolution = VariantBoolProperty(get_index=16, put_index=17)

    GOTO_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p, ctypes.c_long)(7, "Goto")
    START_JOG_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p, ctypes.c_long)(8, "StartJog")
    CHANGE_VELOCITY_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p, ctypes.c_long)(9, "ChangeVelocity")
    STOP_JOG_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_long)(10, "StopJog")
    RESET_POSITION_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_long)(11, "ResetPosition")
    GET_POSITION_RANGE_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p, ctypes.c_void_p)(14, "GetPositionRange")
    CREATE_POSITION_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT)(18, "CreatePosition")
    CREATE_VELOCITY_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT)(19, "CreateVelocity")

    def Goto(self, position, axisMask):
        pos = PiezoStagePosition(position)
        PiezoStage.GOTO_METHOD(self.get(), pos.get(), axisMask)

    def StartJog(self, velocity, axisMask):
        v = PiezoStageVelocity(velocity)
        PiezoStage.START_JOG_METHOD(self.get(), v.get(), axisMask)

    def ChangeVelocity(self, velocity, axisMask):
        v = PiezoStageVelocity(velocity)
        PiezoStage.CHANGE_VELOCITY_METHOD(self.get(), v.get(), axisMask)

    def StopJog(self, axisMask):
        PiezoStage.STOP_JOG_METHOD(self.get(), axisMask)

    def ResetPosition(self, axisMask):
        PiezoStage.RESET_POSITION_METHOD(self.get(), axisMask)

    def GetPositionRange(self):
        pmin, pmax = PiezoStagePosition(), PiezoStagePosition()
        PiezoStage.GET_POSITION_RANGE_METHOD(self.get(), pmin.byref(), pmax.byref())
        return pmin, pmax

    def CreatePosition(self):
        PiezoStage.CREATE_POSITION_METHOD(self.get())

    def CreateVelocity(self):
        PiezoStage.CREATE_VELOCITY_METHOD(self.get())


class UserDoorHatch(IUnknown):
    IID = UUID("2dfdb202-e61d-467e-b59c-c3d21d16c903")

    State = EnumProperty(HatchState, get_index=9)
    IsControlAllowed = VariantBoolProperty(get_index=10)

    OPEN_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT)(7, "Open")
    CLOSE_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT)(8, "Close")

    def Open(self):
        UserDoorHatch.OPEN_METHOD(self.get())

    def Close(self):
        UserDoorHatch.CLOSE_METHOD(self.get())


class FegFlashing(IUnknown):
    IID = UUID("25258776-0e99-4d29-9220-07edec3dbf6d")

    IS_FLASHING_ADVISED_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_int, ctypes.c_void_p)(7, "IsFlashingAdvised")
    PERFORM_FLASHING_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_int)(8, "PerformFlashing")

    def IsFlashingAdvised(self, flashing_type):
        result = ctypes.c_short(-1)
        FegFlashing.IS_FLASHING_ADVISED_METHOD(self.get(), FegFlashingType[flashing_type.upper()], ctypes.byref(result))
        return bool(result.value)

    def PerformFlashing(self, flashing_type):
        FegFlashing.PERFORM_FLASHING_METHOD(self.get(), FegFlashingType[flashing_type.upper()])


class Feg(IUnknown):
    IID = UUID("b2bc347f-c31f-4a51-b6B4-94b9bc2c17fb")

    State = EnumProperty(FegState, get_index=7)
    BeamCurrent = DoubleProperty(get_index=8)
    ExtractorVoltage = DoubleProperty(get_index=9)
    Flashing = ObjectProperty(FegFlashing, get_index=10)
    FocusIndex = FegFocusIndexProperty(get_index=11)


class AdvancedInstrument(IUnknown):
    IID = UUID("c7452318-374a-4161-9b68-90ca3c3f5bea")

    Acquisitions = ObjectProperty(Acquisitions, get_index=7)
    Phaseplate = ObjectProperty(Phaseplate, get_index=8)
    PiezoStage = ObjectProperty(PiezoStage, get_index=9)
    UserDoorHatch = ObjectProperty(UserDoorHatch, get_index=10)
    Source = ObjectProperty(Feg, get_index=11)


class CameraSettingsInternal(IUnknown):
    IID = UUID("e76bb771-821b-4ed2-b7aF-50de5749690f")

    IsTemScriptingLicensed = VariantBoolProperty(get_index=7)
    IsDoseFractionsLicensed = VariantBoolProperty(get_index=8)
    IsElectronCountingLicensed = VariantBoolProperty(get_index=9)


class AdvancedInstrumentInternal(IUnknown):
    IID = UUID("3d50ee81-db96-4c44-a733-ccd0032571e6")

    IsTemScriptingLicensed = VariantBoolProperty(get_index=7)


# TEMAdvancedScripting.AdvancedInstrument.2
CLSID_ADV_INSTRUMENT = UUID("B89721DF-F6F8-4567-9293-D2228012985D")


def GetAdvancedInstrument():
    """Returns Advanced Instrument instance."""
    instrument = co_create_instance(CLSID_ADV_INSTRUMENT, CLSCTX_ALL, AdvancedInstrument)
    return instrument
