import logging
from ctypes import *
from comtypes.client import Constants, CreateObject

BYTE = c_byte
WORD = c_ushort
DWORD = c_ulong

_ole32 = oledll.ole32
_StringFromCLSID = _ole32.StringFromCLSID
_CoTaskMemFree = windll.ole32.CoTaskMemFree
_ProgIDFromCLSID = _ole32.ProgIDFromCLSID
_CLSIDFromString = _ole32.CLSIDFromString
_CLSIDFromProgID = _ole32.CLSIDFromProgID


logging.basicConfig(level=logging.INFO,
                    handlers=[
                        logging.FileHandler("debug.log"),
                        logging.StreamHandler()])

items = [
    ('TEM Scripting', ('Tecnai.Instrument', 'TEMScripting.Instrument.1')),
    ('TOM Moniker', ('TEM.Instrument.1',)),
    ('TEM Advanced Scripting',
     ('TEMAdvancedScripting.AdvancedInstrument.1', 'TEMAdvancedScripting.AdvancedInstrument.2')),
    ('Tecnai Low Dose Kit', ('LDServer.LdSrv',)),
    # ('Gatan CCD Camera', ('TecnaiCCD.GatanCamera.2',)),
    ('TIA', ('ESVision.Application',)),
    ('TVIPS EmMenu', ('EMMENU4.EMMENUApplication.1',)),
]


class GUID(Structure):
    _fields_ = [("Data1", DWORD),
                ("Data2", WORD),
                ("Data3", WORD),
                ("Data4", BYTE * 8)]

    def __init__(self, name=None):
        if name is not None:
            _CLSIDFromString(str(name), byref(self))

    def __repr__(self):
        return 'GUID("%s")' % str(self)

    def __unicode__(self):
        p = c_wchar_p()
        _StringFromCLSID(byref(self), byref(p))
        result = p.value
        _CoTaskMemFree(p)
        return result
    __str__ = __unicode__

    def from_progid(cls, progid):
        """Get guid from progid """
        inst = cls()
        _CLSIDFromProgID(str(progid), byref(inst))
        return inst

    from_progid = classmethod(from_progid)


def run():
    print("Looking for COM module files from type libraries...")
    for item in items:
        message, comnames = item
        for name in comnames:
            try:
                clsid = GUID.from_progid(name)
                logging.info("%s: %s -> %s", message, name, clsid)
            except:
                pass


def run_tree():
    titan = CreateObject("TEMScripting.Instrument.1")
    constants = Constants(titan)

    objs = [
        titan,
        titan.Gun,
        titan.Illumination,
        titan.InstrumentModeControl,
        titan.Projection,
        titan.BlankerShutter,
        titan.Vacuum,
        titan.Stage,
        titan.Stage.Position,
        titan.AutoLoader,
        titan.Acquisition,
        titan.Acquisition.Cameras,
        titan.Acquisition.Cameras[0],
        titan.Acquisition.Cameras[0].Info,
        titan.Acquisition.Cameras[0].AcqParams,
        titan.Camera,
        titan.TemperatureControl,
        titan.UserButtons,
        titan.UserButtons[0],
        titan.Configuration,
        #titan.ApertureMechanismCollection,
    ]

    #run_logs(objs)


def run_tree_adv():
    titan = CreateObject("TEMAdvancedScripting.AdvancedInstrument.2")
    constants = Constants(titan)
    acq = titan.Acquisitions
    csa = acq.CameraSingleAcquisition
    cams = csa.SupportedCameras
    csa.Camera = cams[[c.name for c in cams].index("BM-Falcon")]
    cs = csa.CameraSettings
    dfd = cs.DoseFractionsDefinition
    dfd.Clear()
    dfd.AddRange(1, 2)

    ai = csa.Acquire()
    md = ai.Metadata

    objs = [
        titan,
        titan.Acquisitions,
        titan.Phaseplate,
        acq.CameraSingleAcquisition,
        csa.SupportedCameras,
        csa.SupportedCameras[1],
        cs,
        cs.Binning,
        cs.Capabilities,
        cs.Capabilities.SupportedBinnings,
        cs.Capabilities.SupportedBinnings[0],
        cs.Capabilities.ExposureTimeRange,
        cs.DoseFractionsDefinition,
        cs.DoseFractionsDefinition[0],
        ai,
        md,
        md[0],
    ]

    run_logs(objs)

    logging.info("Falcon metadata:")
    for i in md:
        logging.info("%s --> %s" % (str(i.Key), str(i.ValueAsString)))


def run_logs(objs):
    for obj in objs:
        logging.info(f"{str(obj)} uuid: {str(obj._iid_).lower()}")
        logging.info(f"{str(obj)} methods:")
        for m in obj._methods_:
            logging.info(str(m))
        logging.info("\n\n\n")


if __name__ == '__main__':
    #run()
    run_tree()
    run_tree_adv()


'''
Tecnai.Instrument -> {02CDC9A1-1F1D-11D3-AE11-00A024CBA50C}
TEMScripting.Instrument.1 -> {02CDC9A1-1F1D-11D3-AE11-00A024CBA50C}
TEM.Instrument.1 -> {7D82B1B8-3A42-495A-B1D9-2BE40FA497FA}
TEMAdvancedScripting.AdvancedInstrument.2 -> {B89721DF-F6F8-4567-9293-D2228012985D}
LDServer.LdSrv -> {9BEC9756-A820-11D3-972E-81B6519D0DF8}
ESVision.Application -> {D20B86BB-1214-11D2-AD14-00A0241857FD}

DetectorName --> BM-Falcon
Binning.Width --> 1
Binning.Height --> 1
ReadoutArea.Left --> 0
ReadoutArea.Top --> 0
ReadoutArea.Right --> 4096
ReadoutArea.Bottom --> 4096
ExposureMode --> None
ExposureTime --> 0.99692
DarkGainCorrectionType --> DarkGain
Shutters[0].Type --> Electrostatic
Shutters[0].Position --> PreSpecimen
AcquisitionUnit --> CameraImage
BitsPerPixel --> 32
Encoding --> Signed
ImageSize.Width --> 4096
ImageSize.Height --> 4096
Offset.X --> -4.6106e-07
Offset.Y --> -4.6106e-07
PixelSize.Width --> 2.25127e-10
PixelSize.Height --> 2.25127e-10
PixelUnitX --> m
PixelUnitY --> m
TimeStamp --> 1648565310713477
PixelValueToCameraCounts --> 39
ExposureTime --> 0.971997
CountsToElectrons --> 0.00147737
ElectronCounted --> FALSE
AlignIntegratedImage --> FALSE

'''
