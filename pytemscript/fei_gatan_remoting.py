from .utils.enums import *


class FEIGatanRemoting:
    """ Main class that uses FEI Gatan Remoting from TFS on Gatan PC. """
    def __init__(self, microscope):
        self._plugin = microscope._fei_gatan
        if self._plugin is not None:
            self._plugin_acq = self._plugin.Acquisition
            self._plugin_diag = self._plugin.Diagnostics

    def _find_camera(self, name):
        """Find camera index by name. """
        self._plugin_acq.Initialize()
        for i in range(self._plugin_acq.CameraCount):
            if self._plugin_acq.GetCameraName() == name:
                return i
        raise KeyError("No camera with name %s" % name)

    def acquire_image(self):
        self._plugin_acq.Acquire()
        img = self._plugin.IntegratedImage.GetIntegratedImage  # unsigned char safe array (byte)

    def acquire_movie(self):
        self._plugin_acq.AcquireWithFractions()

    def acquire_dark(self):
        acq_dark = self._plugin_acq.AcquireDarkReferenceImage  # bool rw

    def _set_camera_param(self, name, size, exp_time, binning, **kwargs):
        """ Find the TEM camera and set its params. """
        camera_index = self._find_camera(name)
        print("Found camera ", name, "as #", camera_index)
        print("Location: ", self._plugin_acq.GetCameraLocation(camera_index))  # str
        print("Pixel size:", self._plugin_acq.GetCameraPixelSizeX(camera_index),
              self._plugin_acq.GetCameraPixelSizeY(camera_index))

        self._plugin_acq.SelectCamera(camera_index)
        print("Selected camera #", self._plugin_acq.SelectedCamera)

        if self._plugin_acq.IsCameraRetractable(camera_index):
            self._plugin_acq.RetractCamera(camera_index)
            print("Retracted")
            self._plugin_acq.InsertCamera(camera_index)
            if self._plugin_acq.IsCameraInserted(camera_index):
                print("Inserted")

    def show_params(self):
        exp = self._plugin_acq.ExposureTime  # rw double
        binx, binY = self._plugin_acq.BinningX, self._plugin_acq.BinningY  #rw int

        self._plugin_acq.SetReadoutArea(top=0, left=0, bottom=1, right=1)  # all int
        self._plugin_acq.ResetReadoutArea()

        proc = self._plugin_acq.ProcessingOption  #rw enum AcquisitionProcessingOption
        use_packed = self._plugin_acq.UsePackedData  # rw bool
        frac_time = self._plugin_acq.FractionExposureTime  # rw double
        rmode = self._plugin_acq.ReadMode  # rw enum AcquisitionReadMode
        trans = self._plugin_acq.TransposeSetting  # rw enum TransposeSetting
        store = self._plugin_acq.StoragePath  # rw str
        fn = self._plugin_acq.StorageFilename  # rw str
        fmt = self._plugin_acq.StorageFileFormat  # rw enum FractionFileFormat
        store_dark = self._plugin_acq.StoreDarkReferenceImage  # rw bool
        store_gain = self._plugin_acq.StoreGainReferenceImage  # rw bool
        ok = self._plugin_acq.ValidateStoragePath()  # bool

        space = self._plugin_acq.GetAvailableMegabytesOnStorageLocation()  # long
        revise = self._plugin_acq.Revise()  # bool

        w, h = self._plugin_acq.AcquisitionResultWidth, self._plugin_acq.AcquisitionResultHeight  # long
        pixfmt = self._plugin_acq.AcquisitionResultPixelFormat  # enum PixelFormat
        num_fracs = self._plugin_acq.AcquisitionResultNumberOfFractions  # long
        dose = self._plugin_acq.AcquisitionResultDoseRate  # double

        size = self._plugin_acq.GetDetectorSize(camera_index=0)  #enum?

    def diagnostics(self):
        plugin_version = self._plugin_diag.PluginVersion  # int
        dm_version = self._plugin_diag.DigitalMicrographVersion  # str
        show = self._plugin_diag.ShowAcquiredImagesInDigitalMicrograph  # rw bool

        print(plugin_version, dm_version, show)

        self._plugin_diag.PrintDebug("debug msg")
        self._plugin_diag.PrintResult("result msg")
