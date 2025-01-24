from ..utils.enums import CassetteSlotStatus


class Autoloader:
    """ Sample loading functions. """
    def __init__(self, client):
        self._client = client
        self._has_autoloader_adv = None
        self._err_msg = "Autoloader is not available"
        self._err_msg_adv = "This function is not available in your advanced scripting interface."

    @property
    def __adv_available(self) -> bool:
        if self._has_autoloader_adv is None:
            self._has_autoloader_adv = self._client.has("tem_adv.AutoLoader")
        return self._has_autoloader_adv

    @property
    def is_available(self) -> bool:
        """ Status of the autoloader. Should be always False on Tecnai instruments. """
        return self._client.get("tem.AutoLoader.AutoLoaderAvailable")

    @property
    def number_of_slots(self) -> int:
        """ The number of slots in a cassette. """
        if self.is_available:
            return self._client.get("tem.AutoLoader.NumberOfCassetteSlots")
        else:
            raise RuntimeError(self._err_msg)

    def load_cartridge(self, slot: int) -> None:
        """ Loads the cartridge in the given slot into the microscope.

        :param slot: Slot number
        :type slot: int
        """
        if self.is_available:
            total = self.number_of_slots
            slot = int(slot)
            if slot > total:
                raise ValueError("Only %s slots are available" % total)
            if self.slot_status(slot) != CassetteSlotStatus.OCCUPIED.name:
                raise RuntimeError("Slot %d is not occupied" % slot)
            self._client.call("tem.AutoLoader.LoadCartridge()", slot)
        else:
            raise RuntimeError(self._err_msg)

    def unload_cartridge(self) -> None:
        """ Unloads the cartridge currently in the microscope and puts it back into its
        slot in the cassette.
        """
        if self.is_available:
            self._client.call("tem.AutoLoader.UnloadCartridge()")
        else:
            raise RuntimeError(self._err_msg)

    def run_inventory(self) -> None:
        """ Performs an inventory of the cassette.
        Note: This function takes considerable time to execute.
        """
        # TODO: check if cassette is present
        if self.is_available:
            self._client.call("tem.AutoLoader.PerformCassetteInventory()")
        else:
            raise RuntimeError(self._err_msg)

    def slot_status(self, slot: int) -> str:
        """ The status of the slot specified.

        :param slot: Slot number
        :type slot: int
        """
        if self.is_available:
            total = self.number_of_slots
            if slot > total:
                raise ValueError("Only %s slots are available" % total)
            status = self._client.call("tem.AutoLoader.SlotStatus()", int(slot))
            return CassetteSlotStatus(status).name
        else:
            raise RuntimeError(self._err_msg)

    def undock_cassette(self) -> None:
        """ Moves the cassette from the docker to the capsule. """
        if self.__adv_available:
            if self.is_available:
                self._client.call("tem_adv.AutoLoader.UndockCassette()")
            else:
                raise RuntimeError(self._err_msg)
        else:
            raise NotImplementedError(self._err_msg_adv)

    def dock_cassette(self) -> None:
        """ Moves the cassette from the capsule to the docker. """
        if self.__adv_available:
            if self.is_available:
                self._client.call("tem_adv.AutoLoader.DockCassette()")
            else:
                raise RuntimeError(self._err_msg)
        else:
            raise NotImplementedError(self._err_msg_adv)

    def initialize(self) -> None:
        """ Initializes / Recovers the Autoloader for further use. """
        if self.__adv_available:
            if self.is_available:
                self._client.call("tem_adv.AutoLoader.Initialize()")
            else:
                raise RuntimeError(self._err_msg)
        else:
            raise NotImplementedError(self._err_msg_adv)

    def buffer_cycle(self) -> None:
        """ Synchronously runs the Autoloader buffer cycle. """
        if self.__adv_available:
            if self.is_available:
                self._client.call("tem_adv.AutoLoader.BufferCycle()")
            else:
                raise RuntimeError(self._err_msg)
        else:
            raise NotImplementedError(self._err_msg_adv)
