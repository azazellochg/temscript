from ..utils.enums import HatchState


class UserDoor:
    """ User door hatch controls. Requires advanced scripting. """
    def __init__(self, client):
        self._client = client
        self._err_msg = "Door control is unavailable"
        self._tem_door = None

    @property
    def __adv_available(self) -> bool:
        if self._tem_door is None:
            self._tem_door = self._client.has("tem_adv.UserDoorHatch")
        return self._tem_door

    @property
    def state(self) -> str:
        """ Returns door state. """
        if not self.__adv_available:
            raise NotImplementedError(self._err_msg)
        return HatchState(self._client.get("tem_adv.UserDoorHatch.State")).name

    def open(self) -> None:
        """ Open the door. """
        if self.__adv_available and self._client.get("tem_adv.UserDoorHatch.IsControlAllowed"):
            self._client.call("tem_adv.UserDoorHatch.Open()")
        else:
            raise NotImplementedError(self._err_msg)

    def close(self) -> None:
        """ Close the door. """
        if self.__adv_available and self._client.get("tem_adv.UserDoorHatch.IsControlAllowed"):
            self._client.call("tem_adv.UserDoorHatch.Close()")
        else:
            raise NotImplementedError(self._err_msg)
