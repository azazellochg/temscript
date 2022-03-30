from .enums import *
from ._properties import *


__all__ = ('GetAdvancedInstrument', 'CameraAcquisitionCapabilities', 'CameraSettings',
           'AdvancedCamera', 'AcquiredImage', 'CameraSingleAcquisition',
           'Acquisitions', 'Phaseplate', 'AdvancedInstrument')


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
        range = FrameRange()
        range.Begin, range.End = value[0], value[1]
        FrameRangeList.ADD_METHOD(self.get(), range.get())

    def AddRange(self, begin, end):
        FrameRangeList.ADD_RANGE_METHOD(self.get(), int(begin), int(end))

    def Clear(self):
        FrameRangeList.CLEAR_METHOD(self.get())


class KeyValuePair(IUnknown):
    IID = UUID("565dfad5-a223-44c1-bc09-22c450b21d24")

    Key = StringProperty(get_index=7)
    ValueAsString = StringProperty(get_index=8)


class CameraAcquisitionCapabilities(IUnknown):
    IID = UUID("c4c83905-0d53-47f3-8ae4-91f1248aa0f9")

    _SupportedBinnings = CollectionProperty(get_index=7)
    ExposureTimeRange = ObjectProperty(TimeRange, get_index=8)
    SupportsDoseFractions = VariantBoolProperty(get_index=9)
    MaximumNumberOfDoseFractions = LongProperty(get_index=10)
    SupportsDriftCorrection = VariantBoolProperty(get_index=11)
    SupportsElectronCounting = VariantBoolProperty(get_index=12)
    _SupportsEER = VariantBoolProperty(get_index=13)

    @property
    def SupportedBinnings(self):
        collection = self._SupportedBinnings
        return [Binning(item) for item in collection]

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
    EER = VariantBoolProperty(get_index=23, put_index=24)

    CALCULATE_NUMBER_OF_FRAMES_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT)(22, "CalculateNumberOfFrames")

    def CalculateNumberOfFrames(self):
        CameraSettings.CALCULATE_NUMBER_OF_FRAMES_METHOD(self.get())


class AdvancedCamera(IUnknown):
    IID = UUID("ba60831e-7a06-4f1d-abe9-80d26227fcb9")

    Name = StringProperty(get_index=7)
    Width = LongProperty(get_index=8)
    Height = LongProperty(get_index=9)
    PixelSize = VectorProperty(get_index=10)
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
    _Metadata = CollectionProperty(get_index=11)
    _AsSafeArray = SafeArrayProperty(get_index=12)

    SAVE_TO_FILE_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_wchar_p, ctypes.c_short)(14, "SaveToFile")

    def SaveToFile(self, filePath, normalize=False):
        name_bstr = BStr(filePath)
        bool_value = 0xffff if normalize else 0x0000
        AcquiredImage.SAVE_TO_FILE_METHOD(self.get(), name_bstr.get(), bool_value)

    @property
    def Metadata(self):
        collection = self._Metadata
        return [KeyValuePair(item) for item in collection]

    @property
    def Array(self):
        return self._AsSafeArray.as_array()


class CameraSingleAcquisition(IUnknown):
    IID = UUID("a927ea10-74b0-45a9-8368-fcdd52498053")

    _SupportedCameras = CollectionProperty(get_index=7)
    Camera = ObjectProperty(AdvancedCamera, get_index=None, put_index=8)  # maybe CollectionProperty?
    CameraSettings = ObjectProperty(CameraSettings, get_index=9)
    IsActive = VariantBoolProperty(get_index=11)

    ACQUIRE_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT)(10, "Acquire")
    WAIT_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT)(12, "Wait")

    @property
    def SupportedCameras(self):
        collection = self._SupportedCameras
        return [AdvancedCamera(item) for item in collection]

    @property
    def Acquire(self):
        image = AcquiredImage()
        CameraSingleAcquisition.ACQUIRE_METHOD(self.get(), image.byref())
        return image

    def Wait(self):
        CameraSingleAcquisition.WAIT_METHOD(self.get())


class Acquisitions(IUnknown):
    IID = UUID("27f7ddc7-bad9-4e2e-b9b6-e7644eb152ec")

    _Cameras = CollectionProperty(get_index=7)
    CameraSingleAcquisition = ObjectProperty(CameraSingleAcquisition, get_index=8)

    @property
    def Cameras(self):
        collection = self._Cameras
        return [AdvancedCamera(item) for item in collection]


class Phaseplate(IUnknown):
    IID = UUID("2605f3d9-9365-42fb-8b09-78f8d6b114d4")

    GetCurrentPresetPosition = LongProperty(get_index=8)

    SELECT_NEXT_PRESET_POSITION_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT)(7, "SelectNextPresetPosition")

    def SelectNextPresetPosition(self):
        Phaseplate.SELECT_NEXT_PRESET_POSITION_METHOD(self.get())


class AdvancedInstrument(IUnknown):
    IID = UUID("c7452318-374a-4161-9b68-90ca3c3f5bea")

    Acquisitions = ObjectProperty(Acquisitions, get_index=7)
    Phaseplate = ObjectProperty(Phaseplate, get_index=8)
    PiezoStage = None
    UserDoorHatch = None
    Source = None


# TEMAdvancedScripting.AdvancedInstrument.2
CLSID_ADV_INSTRUMENT = UUID("B89721DF-F6F8-4567-9293-D2228012985D")


def GetAdvancedInstrument():
    """Returns Advanced Instrument instance."""
    instrument = co_create_instance(CLSID_ADV_INSTRUMENT, CLSCTX_ALL, AdvancedInstrument)
    return instrument
