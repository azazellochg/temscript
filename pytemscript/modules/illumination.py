import math
from .utilities import Vector
from ..utils.enums import CondenserLensSystem, CondenserMode, DarkFieldMode, IlluminationMode


class Illumination:
    """ Illumination functions. """
    def __init__(self, client, condenser_type):
        self._client = client
        self._condenser_type = condenser_type
    
    @property
    def __has_3cond(self):
        return self._condenser_type == CondenserLensSystem.THREE_CONDENSER_LENSES.name

    @property
    def spotsize(self):
        """ Spotsize number, usually 1 to 11. (read/write)"""
        return self._client.get("tem.Illumination.SpotsizeIndex")

    @spotsize.setter
    def spotsize(self, value):
        if not (1 <= int(value) <= 11):
            raise ValueError("%s is outside of range 1-11" % value)
        self._client.set("tem.Illumination.SpotsizeIndex", int(value))

    @property
    def intensity(self):
        """ Intensity / C2 condenser lens value. (read/write)"""
        return self._client.get("tem.Illumination.Intensity")

    @intensity.setter
    def intensity(self, value):
        if not (0.0 <= value <= 1.0):
            raise ValueError("%s is outside of range 0.0-1.0" % value)
        self._client.set("tem.Illumination.Intensity", float(value))

    @property
    def intensity_zoom(self):
        """ Intensity zoom. Set to False to disable. (read/write)"""
        return self._client.get("tem.Illumination.IntensityZoomEnabled")

    @intensity_zoom.setter
    def intensity_zoom(self, value):
        self._client.set("tem.Illumination.IntensityZoomEnabled", bool(value))

    @property
    def intensity_limit(self):
        """ Intensity limit. Set to False to disable. (read/write)"""
        return self._client.get("tem.Illumination.IntensityLimitEnabled")

    @intensity_limit.setter
    def intensity_limit(self, value):
        self._client.set("tem.Illumination.IntensityLimitEnabled", bool(value))

    @property
    def beam_shift(self):
        """ Beam shift X and Y in um. (read/write)"""
        return (self._client.get("tem.Illumination.Shift.X") * 1e6,
                self._client.get("tem.Illumination.Shift.Y") * 1e6)

    @beam_shift.setter
    def beam_shift(self, values):
        new_value = Vector(values[0] * 1e-6, values[1] * 1e-6)
        self._client.set("tem.Illumination.Shift", new_value)

    @property
    def rotation_center(self):
        """ Rotation center X and Y in mrad. (read/write)
            Depending on the scripting version,
            the values might need scaling by 6.0 to get mrads.
        """
        return (self._client.get("tem.Illumination.RotationCenter.X") * 1e3,
                self._client.get("tem.Illumination.RotationCenter.Y") * 1e3)

    @rotation_center.setter
    def rotation_center(self, values):
        new_value = Vector(values[0] * 1e-3, values[1] * 1e-3)
        self._client.set("tem.Illumination.RotationCenter", new_value)

    @property
    def condenser_stigmator(self):
        """ C2 condenser stigmator X and Y. (read/write)"""
        return (self._client.get("tem.Illumination.CondenserStigmator.X"),
                self._client.get("tem.Illumination.CondenserStigmator.Y"))

    @condenser_stigmator.setter
    def condenser_stigmator(self, values):
        new_value = Vector(*values)
        new_value.set_limits(-1.0, 1.0)
        self._client.set("tem.Illumination.CondenserStigmator", new_value)

    @property
    def illuminated_area(self):
        """ Illuminated area. Works only on 3-condenser lens systems. (read/write)"""
        if self.__has_3cond:
            return self._client.get("tem.Illumination.IlluminatedArea")
        else:
            raise NotImplementedError("Illuminated area exists only on 3-condenser lens systems.")

    @illuminated_area.setter
    def illuminated_area(self, value):
        if self.__has_3cond:
            self._client.set("tem.Illumination.IlluminatedArea", float(value))
        else:
            raise NotImplementedError("Illuminated area exists only on 3-condenser lens systems.")

    @property
    def probe_defocus(self):
        """ Probe defocus. Works only on 3-condenser lens systems. (read/write)"""
        if self.__has_3cond:
            return self._client.get("tem.Illumination.ProbeDefocus")
        else:
            raise NotImplementedError("Probe defocus exists only on 3-condenser lens systems.")

    @probe_defocus.setter
    def probe_defocus(self, value):
        if self.__has_3cond:
            self._client.set("tem.Illumination.ProbeDefocus", float(value))
        else:
            raise NotImplementedError("Probe defocus exists only on 3-condenser lens systems.")

    #TODO: check if the illum. mode is probe?
    @property
    def convergence_angle(self):
        """ Convergence angle. Works only on 3-condenser lens systems. (read/write)"""
        if self.__has_3cond:
            return self._client.get("tem.Illumination.ConvergenceAngle")
        else:
            raise NotImplementedError("Convergence angle exists only on 3-condenser lens systems.")

    @convergence_angle.setter
    def convergence_angle(self, value):
        if self.__has_3cond:
            self._client.set("tem.Illumination.ConvergenceAngle", float(value))
        else:
            raise NotImplementedError("Convergence angle exists only on 3-condenser lens systems.")

    @property
    def C3ImageDistanceParallelOffset(self):
        """ C3 image distance parallel offset. Works only on 3-condenser lens systems. (read/write)"""
        if self.__has_3cond:
            return self._client.get("tem.Illumination.C3ImageDistanceParallelOffset")
        else:
            raise NotImplementedError("C3ImageDistanceParallelOffset exists only on 3-condenser lens systems.")

    @C3ImageDistanceParallelOffset.setter
    def C3ImageDistanceParallelOffset(self, value):
        if self.__has_3cond:
            self._client.set("tem.Illumination.C3ImageDistanceParallelOffset", float(value))
        else:
            raise NotImplementedError("C3ImageDistanceParallelOffset exists only on 3-condenser lens systems.")

    @property
    def mode(self):
        """ Illumination mode: microprobe or nanoprobe. (read/write)"""
        return IlluminationMode(self._client.get("tem.Illumination.Mode")).name

    @mode.setter
    def mode(self, value):
        self._client.set("tem.Illumination.Mode", value)

    @property
    def dark_field(self):
        """ Dark field mode: cartesian, conical or off. (read/write)"""
        return DarkFieldMode(self._client.get("tem.Illumination.DFMode")).name

    @dark_field.setter
    def dark_field(self, value):
        self._client.set("tem.Illumination.DFMode", value)

    @property
    def condenser_mode(self):
        """ Mode of the illumination system: parallel or probe. (read/write)"""
        if self.__has_3cond:
            return CondenserMode(self._client.get("tem.Illumination.CondenserMode")).name
        else:
            raise NotImplementedError("Condenser mode exists only on 3-condenser lens systems.")

    @condenser_mode.setter
    def condenser_mode(self, value):
        if self.__has_3cond:
            self._client.set("tem.Illumination.CondenserMode", value)
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
        mode = self._client.get("tem.Illumination.DFMode")
        tilt = self._client.get("tem.Illumination.Tilt")
        if mode == DarkFieldMode.CONICAL:
            return tilt[0] * 1e3 * math.cos(tilt[1]), tilt[0] * 1e3 * math.sin(tilt[1])
        elif mode == DarkFieldMode.CARTESIAN:
            return tilt * 1e3
        else:
            return 0.0, 0.0  # Microscope might return nonsense if DFMode is OFF

    @beam_tilt.setter
    def beam_tilt(self, tilt):
        mode = self._client.get("tem.Illumination.DFMode")
        tilt[0] *= 1e-3
        tilt[1] *= 1e-3
        if tilt[0] == 0.0 and tilt[1] == 0.0:
            self._client.set("tem.Illumination.Tilt", (0.0, 0.0))
            self._client.set("tem.Illumination.DFMode", DarkFieldMode.OFF)
        elif mode == DarkFieldMode.CONICAL:
            self._client.set("tem.Illumination.Tilt", (math.sqrt(tilt[0] ** 2 + tilt[1] ** 2), math.atan2(tilt[1], tilt[0])))
        elif mode == DarkFieldMode.OFF:
            self._client.set("tem.Illumination.DFMode", DarkFieldMode.CARTESIAN)
            self._client.set("tem.Illumination.Tilt", tilt)
        else:
            self._client.set("tem.Illumination.Tilt", tilt)
