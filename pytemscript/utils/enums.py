from enum import IntEnum


class TEMScriptingError(IntEnum):
    """ Scripting error codes. """
    E_NOT_OK = -2147155969              # 0x8004ffff
    E_VALUE_CLIP = -2147155970          # 0x8004fffe
    E_OUT_OF_RANGE = -2147155971        # 0x8004fffd
    E_NOT_IMPLEMENTED = -2147155972     # 0x8004fffc
    # The following are also mentioned in the manual
    E_UNEXPECTED = -2147418113          # 0x8000FFFF
    E_NOTIMPL = -2147467263             # 0x80004001
    E_INVALIDARG = -2147024809          # 0x80070057
    E_ABORT = -2147467260               # 0x80004004
    E_FAIL = -2147467259                # 0x80004005
    E_ACCESSDENIED = -2147024891        # 0x80070005


class VacuumStatus(IntEnum):
    """ Vacuum system status. """
    UNKNOWN = 1
    OFF = 2
    CAMERA_AIR = 3
    BUSY = 4
    READY = 5
    ELSE = 6


class GaugeStatus(IntEnum):
    """Vacuum gauge status. """
    UNDEFINED = 0
    UNDERFLOW = 1
    OVERFLOW = 2
    INVALID = 3
    VALID = 4


class GaugePressureLevel(IntEnum):
    """ Vacuum gauge pressure level. """
    UNDEFINED = 0
    LOW = 1
    LOW_MEDIUM = 2
    MEDIUM_HIGH = 3
    HIGH = 4


class StageStatus(IntEnum):
    """ Stage status. """
    READY = 0
    DISABLED = 1
    NOT_READY = 2
    GOING = 3
    MOVING = 4
    WOBBLING = 5


class MeasurementUnitType(IntEnum):
    """ Stage measurement units. """
    UNKNOWN = 0
    METERS = 1
    RADIANS = 2


class StageHolderType(IntEnum):
    """ Specimen holder type. """
    NONE = 0
    SINGLE_TILT = 1
    DOUBLE_TILT = 2
    INVALD = 4
    POLARA = 5
    DUAL_AXIS = 6
    ROTATION_AXIS = 7


class StageAxes(IntEnum):
    """ Stage axes. """
    NONE = 0
    X = 1
    Y = 2
    XY = 3
    Z = 4
    A = 8
    B = 16


class IlluminationNormalization(IntEnum):
    """ Normalization modes for condenser / objective lenses. """
    SPOTSIZE = 1
    INTENSITY = 2
    CONDENSER = 3
    MINI_CONDENSER = 4
    OBJECTIVE = 5
    ALL = 6


class IlluminationMode(IntEnum):
    """ Illumination mode: nanoprobe or microprobe. """
    NANOPROBE = 0
    MICROPROBE = 1


class DarkFieldMode(IntEnum):
    """ Dark field mode. """
    OFF = 1
    CARTESIAN = 2
    CONICAL = 3


class CondenserMode(IntEnum):
    """ Condenser mode: parallel or probe. """
    PARALLEL = 0
    PROBE = 1


class ProjectionNormalization(IntEnum):
    """ Normalization modes for objective/projector lenses. """
    OBJECTIVE = 10
    PROJECTOR = 11
    ALL = 12


class ProjectionMode(IntEnum):
    """ Imaging or diffraction. """
    IMAGING = 1
    DIFFRACTION = 2


class ProjectionSubMode(IntEnum):
    """ Magnification range mode. """
    LM = 1
    M = 2
    SA = 3
    MH = 4
    LAD = 5
    D = 6


class LensProg(IntEnum):
    """ TEM or EFTEM mode. """
    REGULAR = 1
    EFTEM = 2


class ProjectionDetectorShift(IntEnum):
    """ Sets the extra shift that projects the image/diffraction
    pattern onto a detector. """
    ON_AXIS = 0
    NEAR_AXIS = 1
    OFF_AXIS = 2


class ProjDetectorShiftMode(IntEnum):
    """ This property determines whether the chosen DetectorShift
    is changed when the fluorescent screen is moved down. """
    AUTO_IGNORE = 1
    MANUAL = 2
    ALIGNMENT = 3


class HighTensionState(IntEnum):
    """ High Tension status. """
    DISABLED = 1
    OFF = 2
    ON = 3


class InstrumentMode(IntEnum):
    """ TEM or STEM mode. """
    TEM = 0
    STEM = 1


class AcqShutterMode(IntEnum):
    """ Shutter mode. """
    PRE_SPECIMEN = 0
    POST_SPECIMEN = 1
    BOTH = 2


class AcqImageSize(IntEnum):
    """ Image size. """
    FULL = 0
    HALF = 1
    QUARTER = 2


class AcqImageCorrection(IntEnum):
    """ Image correction: unprocessed or corrected (gain/bias). """
    UNPROCESSED = 0
    DEFAULT = 1


class AcqExposureMode(IntEnum):
    """ Exposure mode. """
    NONE = 0
    SIMULTANEOUS = 1
    PRE_EXPOSURE = 2
    PRE_EXPOSURE_PAUSE = 3


class AcqImageFileFormat(IntEnum):
    """ Image file format. """
    TIFF = 0
    JPG = 1
    PNG = 2
    RAW = 3
    SER = 4
    MRC = 5


class ProductFamily(IntEnum):
    """ Microscope product family. """
    TECNAI = 0
    TITAN = 1
    TALOS = 2


class CondenserLensSystem(IntEnum):
    """ Two or three-condenser lens system. """
    TWO_CONDENSER_LENSES = 0
    THREE_CONDENSER_LENSES = 1


class ScreenPosition(IntEnum):
    """ Fluscreen position. """
    UNKNOWN = 1
    UP = 2
    DOWN = 3


class PlateLabelDateFormat(IntEnum):
    """ Date format for film. """
    NO_DATE = 0
    DDMMYY = 1
    MMDDYY = 2
    YYMMDD = 3


class RefrigerantDewar(IntEnum):
    """ Nitrogen dewar. """
    AUTOLOADER_DEWAR = 0
    COLUMN_DEWAR = 1
    HELIUM_DEWAR = 2


class CassetteSlotStatus(IntEnum):
    """ Cassette slot status. """
    UNKNOWN = 0
    OCCUPIED = 1
    EMPTY = 2
    ERROR = 3


class ImagePixelType(IntEnum):
    """ Image type: uint, int or float. """
    UNSIGNED_INT = 0
    SIGNED_INT = 1
    FLOAT = 2


class MechanismId(IntEnum):
    """ Aperture name. """
    UNKNOWN = 0
    C1 = 1
    C2 = 2
    C3 = 3
    OBJ = 4
    SA = 5


class MechanismState(IntEnum):
    """ Aperture state. """
    DISABLED = 0
    INSERTED = 1
    MOVING = 2
    RETRACTED = 3
    ARBITRARY = 4
    HOMING = 5
    ALIGNING = 6
    ERROR = 7


class ApertureType(IntEnum):
    """ Aperture type. """
    UNKNOWN = 0
    CIRCULAR = 1
    BIPRISM = 2
    ENERGY_SLIT = 3
    FARADAY_CUP = 4


class HatchState(IntEnum):
    """ User door hatch state. """
    UNKNOWN = 0
    OPEN = 1
    OPENING = 2
    CLOSED = 3
    CLOSING = 4


class FegState(IntEnum):
    """ FEG state. """
    NOT_EMITTING = 0
    EMITTING = 1


class FegFlashingType(IntEnum):
    """ Cold FEG flashing type. """
    LOW_T = 0
    HIGH_T = 1

# ---------------- Low Dose enums ---------------------------------------------
class LDStatus(IntEnum):
    """ Low Dose status: on or off. """
    IS_OFF = 0
    IS_ON = 1


class LDState(IntEnum):
    """ Low Dose state. """
    SEARCH = 0
    FOCUS1 = 1
    FOCUS2 = 2
    EXPOSURE = 3

# ---------------- FEI Tecnai CCD enums -----------------------------------
class AcqSpeed(IntEnum):
    """ CCD acquisition mode. """
    TURBO = 0
    CONTINUOUS = 1
    SINGLEFRAME = 2


class AcqMode(IntEnum):
    """ CCD acquisition preset."""
    SEARCH = 0
    FOCUS = 1
    RECORD = 2
