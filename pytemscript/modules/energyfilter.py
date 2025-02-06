class EnergyFilter:
    """ Energy filter controls. Requires advanced scripting. """
    def __init__(self, client):
        self._client = client
        self._shortcut = "tem_adv.EnergyFilter"
        self._has_ef = None
        self._err_msg = "EnergyFilter interface is not available"

    @property
    def __adv_available(self) -> bool:
        if self._has_ef is None:
            self._has_ef = self._client.has(self._shortcut)
        return self._has_ef

    def _check_range(self, attrname: str, value: float) -> None:
        vmin = self._client.get(attrname + ".Begin")
        vmax = self._client.get(attrname + ".End")
        if not (vmin <= float(value) <= vmax):
            raise ValueError("Value is outside of allowed "
                             "range: %0.3f - %0.3f" % (vmin, vmax))

    def insert_slit(self, width: float) -> None:
        """ Insert energy slit.

        :param width: Slit width in eV
        :type width: float
        """
        if not self.__adv_available:
            raise NotImplementedError(self._err_msg)
        self._check_range(self._shortcut + ".Slit.WidthRange", width)
        self._client.set(self._shortcut + ".Slit.Width", float(width))
        if not self._client.get(self._shortcut + ".Slit.IsInserted"):
            self._client.call(self._shortcut + ".Slit.Insert()")

    def retract_slit(self) -> None:
        """ Retract energy slit. """
        if not self.__adv_available:
            raise NotImplementedError(self._err_msg)
        self._client.call(self._shortcut + ".Slit.Retract()")

    @property
    def slit_width(self) -> float:
        """ Returns energy slit width in eV. """
        if not self.__adv_available:
            raise NotImplementedError(self._err_msg)
        return self._client.get(self._shortcut + ".Slit.Width")

    @slit_width.setter
    def slit_width(self, value: float) -> None:
        if not self.__adv_available:
            raise NotImplementedError(self._err_msg)
        self._check_range(self._shortcut + ".Slit.WidthRange", value)
        self._client.set(self._shortcut + ".Slit.Width", float(value))

    @property
    def ht_shift(self) -> float:
        """ Returns High Tension energy shift in eV. """
        if not self.__adv_available:
            raise NotImplementedError(self._err_msg)
        return self._client.get(self._shortcut + ".HighTensionEnergyShift.EnergyShift")

    @ht_shift.setter
    def ht_shift(self, value: float) -> None:
        if not self.__adv_available:
            raise NotImplementedError(self._err_msg)
        self._check_range(self._shortcut + ".HighTensionEnergyShift.EnergyShiftRange", value)
        self._client.set(self._shortcut + ".HighTensionEnergyShift.EnergyShift", float(value))

    @property
    def zlp_shift(self) -> float:
        """ Returns Zero-Loss Peak (ZLP) energy shift in eV. """
        if not self.__adv_available:
            raise NotImplementedError(self._err_msg)
        return self._client.get(self._shortcut + ".ZeroLossPeakAdjustment.EnergyShift")

    @zlp_shift.setter
    def zlp_shift(self, value: float) -> None:
        if not self.__adv_available:
            raise NotImplementedError(self._err_msg)
        self._check_range(self._shortcut + ".ZeroLossPeakAdjustment.EnergyShiftRange", value)
        self._client.set(self._shortcut + ".ZeroLossPeakAdjustment.EnergyShift", float(value))
