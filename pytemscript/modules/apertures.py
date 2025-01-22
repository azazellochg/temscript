from ..utils.enums import MechanismId, MechanismState


class Apertures:
    """ Apertures and VPP controls. """
    def __init__(self, client):
        self._client = client
        self._has_apertures = None
        self._err_msg = "Apertures interface is not available. Requires a separate license"
        self._err_msg_vpp = "Either no VPP found or it's not enabled and inserted"

    @property
    def __std_available(self):
        if self._has_apertures is None:
            self._has_apertures = self._client.has("tem.ApertureMechanismCollection")
        return self._has_apertures

    def _find_aperture(self, name):
        """Find aperture object by name. """
        if not self.__std_available:
            raise NotImplementedError(self._err_msg)
        for ap in self._client.get("tem.ApertureMechanismCollection"):
            if MechanismId(ap.Id).name == name.upper():
                return ap
        raise KeyError("No aperture with name %s" % name)

    @property
    def vpp_position(self):
        """ Returns the index of the current VPP preset position. """
        try:
            return self._client.get("tem_adv.PhasePlate.GetCurrentPresetPosition") + 1
        except:
            raise RuntimeError(self._err_msg_vpp)

    def vpp_next_position(self):
        """ Goes to the next preset location on the VPP aperture. """
        try:
            self._client.call("tem_adv.PhasePlate.SelectNextPresetPosition()")
        except:
            raise RuntimeError(self._err_msg_vpp)

    def enable(self, aperture):
        ap = self._find_aperture(aperture)
        ap.Enable()

    def disable(self, aperture):
        ap = self._find_aperture(aperture)
        ap.Disable()

    def retract(self, aperture):
        ap = self._find_aperture(aperture)
        if ap.IsRetractable:
            ap.Retract()

    def select(self, aperture, size):
        """ Select a specific aperture.

        :param aperture: Aperture name (C1, C2, C3, OBJ or SA)
        :type aperture: str
        :param size: Aperture size
        :type size: float
        """
        ap = self._find_aperture(aperture)
        if ap.State == MechanismState.DISABLED:
            ap.Enable()
        for a in ap.ApertureCollection:
            if a.Diameter == size:
                ap.SelectAperture(a)
                if ap.SelectedAperture.Diameter == size:
                    return
                else:
                    raise RuntimeError("Could not select aperture!")

    @property
    def show_all(self):
        """ Returns a dict with apertures information. """
        if not self.__std_available:
            raise NotImplementedError(self._err_msg)
        result = {}
        for ap in self._client.get("tem.ApertureMechanismCollection"):
            result[MechanismId(ap.Id).name] = {"retractable": ap.IsRetractable,
                                               "state": MechanismState(ap.State).name,
                                               "sizes": [a.Diameter for a in ap.ApertureCollection]
                                               }
        return result
