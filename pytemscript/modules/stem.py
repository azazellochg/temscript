from .extras import Vector
from ..utils.enums import InstrumentMode


class Stem:
    """ STEM functions. """
    def __init__(self, client):
        self._client = client
        self._err_msg = "Microscope not in STEM mode"

    @property
    def is_available(self) -> bool:
        """ Returns whether the microscope has a STEM system or not. """
        return self._client.has("tem.InstrumentModeControl.StemAvailable")

    def enable(self) -> None:
        """ Switch to STEM mode."""
        if self.is_available:
            self._client.set("tem.InstrumentModeControl.InstrumentMode", InstrumentMode.STEM)
        else:
            raise RuntimeError(self._err_msg)

    def disable(self) -> None:
        """ Switch back to TEM mode. """
        self._client.set("tem.InstrumentModeControl.InstrumentMode", InstrumentMode.TEM)

    @property
    def magnification(self) -> float:
        """ The magnification value in STEM mode. (read/write)"""
        if self._client.get("tem.InstrumentModeControl.InstrumentMode") == InstrumentMode.STEM:
            return self._client.get("tem.Illumination.StemMagnification")
        else:
            raise RuntimeError(self._err_msg)

    @magnification.setter
    def magnification(self, mag: int) -> None:
        if self._client.get("tem.InstrumentModeControl.InstrumentMode") == InstrumentMode.STEM:
            self._client.set("tem.Illumination.StemMagnification", float(mag))
        else:
            raise RuntimeError(self._err_msg)

    @property
    def rotation(self) -> float:
        """ The STEM rotation angle (in mrad). (read/write)"""
        if self._client.get("tem.InstrumentModeControl.InstrumentMode") == InstrumentMode.STEM:
            return self._client.get("tem.Illumination.StemRotation") * 1e3
        else:
            raise RuntimeError(self._err_msg)

    @rotation.setter
    def rotation(self, rot: float) -> None:
        if self._client.get("tem.InstrumentModeControl.InstrumentMode") == InstrumentMode.STEM:
            self._client.set("tem.Illumination.StemRotation", float(rot) * 1e-3)
        else:
            raise RuntimeError(self._err_msg)

    @property
    def scan_field_of_view(self) -> tuple:
        """ STEM full scan field of view. (read/write)"""
        if self._client.get("tem.InstrumentModeControl.InstrumentMode") == InstrumentMode.STEM:
            return (self._client.get("tem.Illumination.StemFullScanFieldOfView.X"),
                    self._client.get("tem.Illumination.StemFullScanFieldOfView.Y"))
        else:
            raise RuntimeError(self._err_msg)

    @scan_field_of_view.setter
    def scan_field_of_view(self, values: tuple) -> None:
        if self._client.get("tem.InstrumentModeControl.InstrumentMode") == InstrumentMode.STEM:
            new_value = Vector(*values)
            self._client.set("tem.Illumination.StemFullScanFieldOfView", new_value)
        else:
            raise RuntimeError(self._err_msg)
