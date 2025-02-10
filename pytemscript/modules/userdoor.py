from ..utils.enums import HatchState


class UserDoor:
    """ User door hatch controls. Requires advanced scripting. """
    def __init__(self, client):
        self._client = client
        self._shortcut = "tem_adv.UserDoorHatch"
        self._err_msg = "Door control is unavailable"
        self._tem_door = None

    @property
    def __adv_available(self) -> bool:
        if self._tem_door is None:
            self._tem_door = self._client.has(self._shortcut)
        return self._tem_door

    @property
    def state(self) -> str:
        """ Returns door state. """
        if not self.__adv_available:
            raise NotImplementedError(self._err_msg)
        return HatchState(self._client.get(self._shortcut + ".State")).name

    def open(self) -> None:
        """ Open the door. """
        if self.__adv_available and self._client.get(self._shortcut + ".IsControlAllowed"):
            self._client.call(self._shortcut + ".Open()")
        else:
            raise NotImplementedError(self._err_msg)

    def close(self) -> None:
        """ Close the door. """
        if self.__adv_available and self._client.get(self._shortcut + ".IsControlAllowed"):
            self._client.call(self._shortcut + ".Close()")
        else:
            raise NotImplementedError(self._err_msg)
