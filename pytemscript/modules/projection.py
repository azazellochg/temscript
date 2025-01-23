from ..utils.enums import ProjectionMode, ProjectionSubMode, ProjDetectorShiftMode, ProjectionDetectorShift, LensProg
from .utilities import Vector


class Projection:
    """ Projection system functions. """
    def __init__(self, client):
        self._client = client
        self._err_msg = "Microscope is not in diffraction mode"
        #self.magnification_index = self._tem_projection.MagnificationIndex
        #self.camera_length_index = self._tem_projection.CameraLengthIndex

    @property
    def focus(self):
        """ Absolute focus value. (read/write)"""
        return self._client.get("tem.Projection.Focus")

    @focus.setter
    def focus(self, value):
        if not (-1.0 <= value <= 1.0):
            raise ValueError("%s is outside of range -1.0 to 1.0" % value)

        self._client.set("tem.Projection.Focus", float(value))

    @property
    def magnification(self):
        """ The reference magnification value (screen up setting)."""
        if self._client.get("tem.Projection.Mode") == ProjectionMode.IMAGING:
            return self._client.get("tem.Projection.Magnification")
        else:
            raise RuntimeError(self._err_msg)

    @property
    def camera_length(self):
        """ The reference camera length in m (screen up setting). """
        if self._client.get("tem.Projection.Mode") == ProjectionMode.DIFFRACTION:
            return self._client.get("tem.Projection.CameraLength")
        else:
            raise RuntimeError(self._err_msg)

    @property
    def image_shift(self):
        """ Image shift in um. (read/write)"""
        return (self._client.get("tem.Projection.ImageShift.X") * 1e6,
                self._client.get("tem.Projection.ImageShift.Y") * 1e6)

    @image_shift.setter
    def image_shift(self, values):
        new_value = Vector(values[0] * 1e-6, values[1] * 1e-6)
        self._client.set("tem.Projection.ImageShift", new_value)

    @property
    def image_beam_shift(self):
        """ Image shift with beam shift compensation in um. (read/write)"""
        return (self._client.get("tem.Projection.ImageBeamShift.X") * 1e6,
                self._client.get("tem.Projection.ImageBeamShift.Y") * 1e6)

    @image_beam_shift.setter
    def image_beam_shift(self, values):
        new_value = Vector(values[0] * 1e-6, values[1] * 1e-6)
        self._client.set("tem.Projection.ImageBeamShift", new_value)

    @property
    def image_beam_tilt(self):
        """ Beam tilt with diffraction shift compensation in mrad. (read/write)"""
        return (self._client.get("tem.Projection.ImageBeamTilt.X") * 1e3,
                self._client.get("tem.Projection.ImageBeamTilt.Y") * 1e3)

    @image_beam_tilt.setter
    def image_beam_tilt(self, values):
        new_value = Vector(values[0] * 1e-3, values[1] * 1e-3)
        self._client.set("tem.Projection.ImageBeamTilt", new_value)

    @property
    def diffraction_shift(self):
        """ Diffraction shift in mrad. (read/write)"""
        #TODO: 180/pi*value = approx number in TUI
        return (self._client.get("tem.Projection.DiffractionShift.X") * 1e3,
                self._client.get("tem.Projection.DiffractionShift.Y") * 1e3)

    @diffraction_shift.setter
    def diffraction_shift(self, values):
        new_value = Vector(values[0] * 1e-3, values[1] * 1e-3)
        self._client.set("tem.Projection.DiffractionShift", new_value)

    @property
    def diffraction_stigmator(self):
        """ Diffraction stigmator. (read/write)"""
        if self._client.get("tem.Projection.Mode") == ProjectionMode.DIFFRACTION:
            return (self._client.get("tem.Projection.DiffractionStigmator.X"),
                    self._client.get("tem.Projection.DiffractionStigmator.Y"))
        else:
            raise RuntimeError(self._err_msg)

    @diffraction_stigmator.setter
    def diffraction_stigmator(self, values):
        if self._client.get("tem.Projection.Mode") == ProjectionMode.DIFFRACTION:
            new_value = Vector(*values)
            new_value.set_limits(-1.0, 1.0)
            self._client.set("tem.Projection.DiffractionStigmator", new_value)
        else:
            raise RuntimeError(self._err_msg)

    @property
    def objective_stigmator(self):
        """ Objective stigmator. (read/write)"""
        return (self._client.get("tem.Projection.ObjectiveStigmator.X"),
                self._client.get("tem.Projection.ObjectiveStigmator.Y"))

    @objective_stigmator.setter
    def objective_stigmator(self, values):
        new_value = Vector(*values)
        new_value.set_limits(-1.0, 1.0)
        self._client.set("tem.Projection.ObjectiveStigmator", new_value)

    @property
    def defocus(self):
        """ Defocus value in um. (read/write)"""
        return self._client.get("tem.Projection.Defocus") * 1e6

    @defocus.setter
    def defocus(self, value):
        self._client.set("tem.Projection.Defocus", float(value) * 1e-6)

    @property
    def mode(self):
        """ Main mode of the projection system (either imaging or diffraction). (read/write)"""
        return ProjectionMode(self._client.get("tem.Projection.Mode")).name

    @mode.setter
    def mode(self, mode):
        self._client.set("tem.Projection.Mode", mode)

    @property
    def detector_shift(self):
        """ Detector shift. (read/write)"""
        return ProjectionDetectorShift(self._client.get("tem.Projection.DetectorShift")).name

    @detector_shift.setter
    def detector_shift(self, value):
        self._client.set("tem.Projection.DetectorShift", value)

    @property
    def detector_shift_mode(self):
        """ Detector shift mode. (read/write)"""
        return ProjDetectorShiftMode(self._client.get("tem.Projection.DetectorShiftMode")).name

    @detector_shift_mode.setter
    def detector_shift_mode(self, value):
        self._client.set("tem.Projection.DetectorShiftMode", value)

    @property
    def magnification_range(self):
        """ Submode of the projection system (either LM, M, SA, MH, LAD or D).
        The imaging submode can change when the magnification is changed.
        """
        return ProjectionSubMode(self._client.get("tem.Projection.SubMode")).name

    @property
    def image_rotation(self):
        """ The rotation of the image or diffraction pattern on the
        fluorescent screen with respect to the specimen. Units: mrad.
        """
        return self._client.get("tem.Projection.ImageRotation") * 1e3

    @property
    def is_eftem_on(self):
        """ Check if the EFTEM lens program setting is ON. """
        return LensProg(self._client.get("tem.Projection.LensProgram")) == LensProg.EFTEM

    def eftem_on(self):
        """ Switch on EFTEM. """
        self._client.set("tem.Projection.LensProgram", LensProg.EFTEM)

    def eftem_off(self):
        """ Switch off EFTEM. """
        self._client.set("tem.Projection.LensProgram", LensProg.REGULAR)

    def reset_defocus(self):
        """ Reset defocus value in the TEM user interface to zero.
        Does not change any lenses. """
        self._client.call("tem.Projection.ResetDefocus()")
