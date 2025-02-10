from ..utils.enums import LDState, LDStatus


class LowDose:
    """ Low Dose functions. """
    def __init__(self, client):
        self._client = client
        self._err_msg = "Low Dose is not available"

    @property
    def is_available(self) -> bool:
        """ Return True if Low Dose is available. """
        return (self._client.has_lowdose_iface and
                self._client.get("tem_lowdose.LowDoseAvailable") and
                self._client.get("tem_lowdose.IsInitialized"))

    @property
    def is_active(self) -> bool:
        """ Check if the Low Dose is ON. """
        if self.is_available:
            return LDStatus(self._client.get("tem_lowdose.LowDoseActive")) == LDStatus.IS_ON
        else:
            raise RuntimeError(self._err_msg)

    @property
    def state(self) -> str:
        """ Low Dose state (LDState enum). (read/write) """
        if self.is_available and self.is_active:
            return LDState(self._client.get("tem_lowdose.LowDoseState")).name
        else:
            raise RuntimeError(self._err_msg)

    @state.setter
    def state(self, state: LDState) -> None:
        if self.is_available:
            self._client.set("tem_lowdose.LowDoseState", state)
        else:
            raise RuntimeError(self._err_msg)

    def on(self) -> None:
        """ Switch ON Low Dose."""
        if self.is_available:
            self._client.set("tem_lowdose.LowDoseActive", LDStatus.IS_ON)
        else:
            raise RuntimeError(self._err_msg)

    def off(self) -> None:
        """ Switch OFF Low Dose."""
        if self.is_available:
            self._client.set("tem_lowdose.LowDoseActive", LDStatus.IS_OFF)
        else:
            raise RuntimeError(self._err_msg)
