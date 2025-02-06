from typing import Dict

from .extras import StagePosition


class PiezoStage:
    """ Piezo stage functions. """
    def __init__(self, client):
        self._client = client
        self._shortcut = "tem_adv.PiezoStage"
        self._has_pstage = None
        self._err_msg = "PiezoStage interface is not available."

    @property
    def __adv_available(self) -> bool:
        if self._has_pstage is None:
            self._has_pstage = self._client.has(self._shortcut + ".HighResolution")
        return self._has_pstage

    @property
    def position(self) -> Dict:
        """ The current position of the piezo stage (x,y,z in um). """
        if not self.__adv_available:
            raise NotImplementedError(self._err_msg)
        else:
            return self._client.call(self._shortcut + ".CurrentPosition",
                                     obj=StagePosition, func="get")

    @property
    def position_range(self) -> tuple[float, float]:
        """ Return min and max positions. """
        if not self.__adv_available:
            raise NotImplementedError(self._err_msg)
        else:
            return self._client.call(self._shortcut + ".GetPositionRange()")

    @property
    def velocity(self) -> Dict:
        """ Returns a dict with stage velocities. """
        if not self.__adv_available:
            raise NotImplementedError(self._err_msg)
        else:
            return self._client.call(self._shortcut + ".CurrentJogVelocity",
                                     obj=StagePosition, func="get")
