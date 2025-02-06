import logging
import time

from ..utils.enums import FegState, HighTensionState, FegFlashingType
from .extras import Vector


class Gun:
    """ Gun functions. """
    def __init__(self, client):
        self._client = client
        self._has_gun1 = None
        self._has_source = None
        self._err_msg_gun1 = "Gun1 interface is not available. Requires TEM Server 7.10+"
        self._err_msg_cfeg = "Source/C-FEG interface is not available"

    @property
    def __gun1_available(self) -> bool:
        if self._has_gun1 is None:
            self._has_gun1 = self._client.has("tem.Gun1")
        return self._has_gun1

    @property
    def __adv_available(self) -> bool:
        if self._has_source is None:
            self._has_source = self._client.has("tem_adv.Source.State")
        return self._has_source

    @property
    def shift(self) -> tuple:
        """ Gun shift. (read/write)"""
        return (self._client.get("tem.Gun.Shift.X"),
                self._client.get("tem.Gun.Shift.Y"))

    @shift.setter
    def shift(self, values: tuple) -> None:
        new_value = Vector(*values)
        new_value.set_limits(-1.0, 1.0)
        self._client.set("tem.Gun.Shift", new_value)

    @property
    def tilt(self) -> tuple:
        """ Gun tilt. (read/write)"""
        return (self._client.get("tem.Gun.Tilt.X"),
                self._client.get("tem.Gun.Tilt.Y"))

    @tilt.setter
    def tilt(self, values: tuple) -> None:
        new_value = Vector(*values)
        new_value.set_limits(-1.0, 1.0)
        self._client.set("tem.Gun.Tilt", new_value)

    @property
    def voltage_offset(self) -> float:
        """ High voltage offset. (read/write)"""
        if self.__gun1_available:
            return self._client.get("tem.Gun1.HighVoltageOffset")
        else:
            raise NotImplementedError(self._err_msg_gun1)

    @voltage_offset.setter
    def voltage_offset(self, offset: float) -> None:
        if self.__gun1_available:
            self._client.set("tem.Gun1.HighVoltageOffset", float(offset))
        else:
            raise NotImplementedError(self._err_msg_gun1)

    @property
    def feg_state(self) -> str:
        """ FEG emitter status. """
        if self.__adv_available:
            return FegState(self._client.get("tem_adv.Source.State")).name
        else:
            raise NotImplementedError(self._err_msg_cfeg)

    @property
    def ht_state(self) -> str:
        """ High tension state: on, off or disabled.
        Disabling/enabling can only be done via the button on the
        system on/off-panel, not via script. When switching on
        the high tension, this function cannot check if and
        when the set value is actually reached. (read/write)
        """
        return HighTensionState(self._client.get("tem.Gun.HTState")).name

    @ht_state.setter
    def ht_state(self, value: HighTensionState) -> None:
        self._client.set("tem.Gun.HTState", value)

    @property
    def voltage(self) -> float:
        """ The value of the HT setting as displayed in the TEM user
        interface. Units: kVolts. (read/write)
        """
        state = self._client.get("tem.Gun.HTState")
        if state == HighTensionState.ON:
            return self._client.get("tem.Gun.HTValue") * 1e-3
        else:
            return 0.0

    @voltage.setter
    def voltage(self, value: float) -> None:
        voltage_max = self.voltage_max
        if not (0.0 <= value <= voltage_max):
            raise ValueError("%s is outside of range 0.0-%s" % (value, voltage_max))
        self._client.set("tem.Gun.HTValue", float(value) * 1000)
        while True:
            if self._client.get("tem.Gun.HTValue") == float(value) * 1000:
                logging.info("Changing HT voltage complete.")
                break
            else:
                time.sleep(10)

    @property
    def voltage_max(self) -> float:
        """ The maximum possible value of the HT on this microscope. Units: kVolts. """
        return self._client.get("tem.Gun.HTMaxValue") * 1e-3

    @property
    def voltage_offset_range(self):
        """ Returns the high voltage offset range. """
        if self.__gun1_available:
            #TODO: this is a function?
            return self._client.call("tem.Gun1.GetHighVoltageOffsetRange()")
        else:
            raise NotImplementedError(self._err_msg_gun1)

    @property
    def beam_current(self) -> float:
        """ Returns the C-FEG beam current in Amperes. """
        if self.__adv_available:
            return self._client.get("tem_adv.Source.BeamCurrent")
        else:
            raise NotImplementedError(self._err_msg_cfeg)

    @property
    def extractor_voltage(self) -> float:
        """ Returns the extractor voltage. """
        if self.__adv_available:
            return self._client.get("tem_adv.Source.ExtractorVoltage")
        else:
            raise NotImplementedError(self._err_msg_cfeg)

    @property
    def focus_index(self) -> tuple:
        """ Returns coarse and fine gun lens index. """
        if self.__adv_available:
            return (self._client.get("tem_adv.Source.FocusIndex.Coarse"),
                    self._client.get("tem_adv.Source.FocusIndex.Fine"))
        else:
            raise NotImplementedError(self._err_msg_cfeg)

    def do_flashing(self, flash_type: FegFlashingType) -> None:
        """ Perform cold FEG flashing.

        :param flash_type: FEG flashing type (FegFlashingType enum)
        :type flash_type: IntEnum
        """
        if not self.__adv_available:
            raise NotImplementedError(self._err_msg_cfeg)
        if self._client.call("tem_adv.Source.Flashing.IsFlashingAdvised()", flash_type):
            # FIXME: lowT flashing can be done even if not advised
            self._client.call("tem_adv.Source.Flashing.PerformFlashing()", flash_type)
        else:
            raise Warning("Flashing type %s is not advised" % flash_type)
