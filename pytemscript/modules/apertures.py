from typing import Optional, Dict

from .extras import SpecialObj
from ..utils.enums import MechanismId, MechanismState


class AperturesObj(SpecialObj):
    """ Wrapper around apertures COM object. """

    def show(self) -> Dict:
        """ Returns a dict with apertures information. """
        apertures = {}
        for ap in self.com_object:
            apertures[MechanismId(ap.Id).name] = {
                "retractable": ap.IsRetractable,
                "state": MechanismState(ap.State).name,
                "sizes": [a.Diameter for a in ap.ApertureCollection]
            }

        return apertures

    def _find_aperture(self, name: str):
        """ Helper method to find the aperture object by name. """
        name = name.upper()
        for ap in self.com_object:
            if name == MechanismId(ap.Id).name:
                return ap
        raise KeyError("No aperture with name %s" % name)

    def enable(self, name: str) -> None:
        ap = self._find_aperture(name)
        ap.Enable()

    def disable(self, name: str) -> None:
        ap = self._find_aperture(name)
        ap.Disable()

    def retract(self, name: str) -> None:
        ap = self._find_aperture(name)
        if ap.IsRetractable:
            ap.Retract()
        else:
            raise NotImplementedError("Aperture %s is not retractable" % name)

    def select(self, name: str, size: int) -> None:
        ap = self._find_aperture(name)
        if ap.State == MechanismState.DISABLED:
            ap.Enable()
        for a in ap.ApertureCollection:
            if a.Diameter == size:
                ap.SelectAperture(a)
                if ap.SelectedAperture.Diameter == size:
                    return
                else:
                    raise RuntimeError("Could not select aperture!")


class Apertures:
    """ Apertures and VPP controls. """
    def __init__(self, client):
        self._client = client
        self._has_apertures = None
        self._shortcut = "tem.ApertureMechanismCollection"
        self._err_msg = "Apertures interface is not available. Requires a separate license"
        self._err_msg_vpp = "Either no VPP found or it's not enabled and inserted"

    @property
    def __std_available(self) -> bool:
        if self._has_apertures is None:
            self._has_apertures = self._client.has(self._shortcut)
        return self._has_apertures

    @property
    def vpp_position(self) -> int:
        """ Returns the index of the current VPP preset position. """
        try:
            return int(self._client.get("tem_adv.PhasePlate.GetCurrentPresetPosition")) + 1
        except:
            raise RuntimeError(self._err_msg_vpp)

    def vpp_next_position(self) -> None:
        """ Goes to the next preset location on the VPP aperture. """
        try:
            self._client.call("tem_adv.PhasePlate.SelectNextPresetPosition()")
        except:
            raise RuntimeError(self._err_msg_vpp)

    def enable(self, aperture) -> None:
        if not self.__std_available:
            raise NotImplementedError(self._err_msg)
        else:
            self._client.call(self._shortcut, obj=AperturesObj,
                              func="enable", name=aperture)

    def disable(self, aperture) -> None:
        if not self.__std_available:
            raise NotImplementedError(self._err_msg)
        else:
            self._client.call(self._shortcut, obj=AperturesObj,
                              func="disable", name=aperture)

    def retract(self, aperture) -> None:
        if not self.__std_available:
            raise NotImplementedError(self._err_msg)
        else:
            self._client.call(self._shortcut, obj=AperturesObj,
                              func="retract", name=aperture)

    def select(self, aperture: str, size: int) -> None:
        """ Select a specific aperture.

        :param aperture: Aperture name (C1, C2, C3, OBJ or SA)
        :type aperture: str
        :param size: Aperture size
        :type size: float
        """
        if not self.__std_available:
            raise NotImplementedError(self._err_msg)
        else:
            self._client.call(self._shortcut, obj=AperturesObj,
                              func="select", name=aperture, size=size)

    def show(self) -> Optional[dict]:
        """ Returns a dict with apertures information. """
        if not self.__std_available:
            raise NotImplementedError(self._err_msg)
        else:
            self._client.call(self._shortcut, obj=AperturesObj,
                              func="show")
