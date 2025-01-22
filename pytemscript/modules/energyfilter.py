
class EnergyFilter:
    """ Energy filter controls. Requires advanced scripting. """
    def __init__(self, client):
        self._client = client
        self._has_ef = None
        self._err_msg = "EnergyFilter interface is not available"

    @property
    def __adv_available(self):
        if self._has_ef is None:
            self._has_ef = self._client.has("tem_adv.EnergyFilter")
        return self._has_ef

    def _check_range(self, attrname, value):
        vmin = self._client.get(attrname + ".Begin")
        vmax = self._client.get(attrname + ".End")
        if not (vmin <= value <= vmax):
            raise ValueError("Value is outside of allowed "
                             "range: %0.3f - %0.3f" % (vmin, vmax))

    def insert_slit(self, width):
        """ Insert energy slit.

        :param width: Slit width in eV
        :type width: float
        """
        if not self.__adv_available:
            raise NotImplementedError(self._err_msg)
        self._check_range("tem_adv.EnergyFilter.Slit.WidthRange", width)
        self._client.set("tem_adv.EnergyFilter.Slit.Width", width)
        if not self._client.get("tem_adv.EnergyFilter.Slit.IsInserted"):
            self._client.call("tem_adv.EnergyFilter.Slit.Insert()")

    def retract_slit(self):
        """ Retract energy slit. """
        if not self.__adv_available:
            raise NotImplementedError(self._err_msg)
        self._client.call("tem_adv.EnergyFilter.Slit.Retract()")

    @property
    def slit_width(self):
        """ Returns energy slit width in eV. """
        if not self.__adv_available:
            raise NotImplementedError(self._err_msg)
        return self._client.get("tem_adv.EnergyFilter.Slit.Width")

    @slit_width.setter
    def slit_width(self, value):
        if not self.__adv_available:
            raise NotImplementedError(self._err_msg)
        self._check_range("tem_adv.EnergyFilter.Slit.WidthRange", value)
        self._client.set("tem_adv.EnergyFilter.Slit.Width", value)

    @property
    def ht_shift(self):
        """ Returns High Tension energy shift in eV. """
        if not self.__adv_available:
            raise NotImplementedError(self._err_msg)
        return self._client.get("tem_adv.EnergyFilter.HighTensionEnergyShift.EnergyShift")

    @ht_shift.setter
    def ht_shift(self, value):
        if not self.__adv_available:
            raise NotImplementedError(self._err_msg)
        self._check_range("tem_adv.EnergyFilter.HighTensionEnergyShift.EnergyShiftRange", value)
        self._client.set("tem_adv.EnergyFilter.HighTensionEnergyShift.EnergyShift", value)

    @property
    def zlp_shift(self):
        """ Returns Zero-Loss Peak (ZLP) energy shift in eV. """
        if not self.__adv_available:
            raise NotImplementedError(self._err_msg)
        return self._client.get("tem_adv.EnergyFilter.ZeroLossPeakAdjustment.EnergyShift")

    @zlp_shift.setter
    def zlp_shift(self, value):
        if not self.__adv_available:
            raise NotImplementedError(self._err_msg)
        self._check_range("tem_adv.EnergyFilter.ZeroLossPeakAdjustment.EnergyShiftRange", value)
        self._client.set("tem_adv.EnergyFilter.ZeroLossPeakAdjustment.EnergyShift", value)
