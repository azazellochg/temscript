class PiezoStage:
    """ Piezo stage functions. """
    def __init__(self, client):
        self._client = client
        self._has_pstage = None
        self._err_msg = "PiezoStage interface is not available."

    @property
    def __adv_available(self) -> bool:
        if self._has_pstage is None:
            self._has_pstage = self._client.has("tem_adv.PiezoStage.HighResolution")
        return self._has_pstage

    @property
    def position(self) -> dict:
        """ The current position of the piezo stage (x,y,z in um). """
        if not self.__adv_available:
            raise NotImplementedError(self._err_msg)
        pos = self._client.get("tem_adv.PiezoStage.CurrentPosition")
        return {key: getattr(pos, key.upper()) * 1e6 for key in 'xyz'}

    @property
    def position_range(self):
        """ Return min and max positions. """
        if not self.__adv_available:
            raise NotImplementedError(self._err_msg)
        return self._client.call("tem_adv.PiezoStage.GetPositionRange()")

    @property
    def velocity(self) -> dict:
        """ Returns a dict with stage velocities. """
        if not self.__adv_available:
            raise NotImplementedError(self._err_msg)
        pos = self._client.get("tem_adv.PiezoStage.CurrentJogVelocity")
        return {key: getattr(pos, key.upper()) for key in 'xyz'}
