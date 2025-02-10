from typing import Dict

from ..utils.enums import VacuumStatus, GaugeStatus, GaugePressureLevel
from .extras import SpecialObj


class GaugesObj(SpecialObj):
    """ Wrapper around vacuum gauges COM object. """

    def show(self) -> Dict:
        """ Returns a dict with vacuum gauges information. """
        gauges = {}
        for g in self.com_object:
            # g.Read()
            if g.Status == GaugeStatus.UNDEFINED:
                # set manually if undefined, otherwise fails
                pressure_level = GaugePressureLevel.UNDEFINED.name
            else:
                pressure_level = GaugePressureLevel(g.PressureLevel).name

            gauges[g.Name] = {
                "status": GaugeStatus(g.Status).name,
                "pressure": g.Pressure,
                "trip_level": pressure_level
            }

        return gauges


class Vacuum:
    """ Vacuum functions. """
    def __init__(self, client):
        self._client = client
        self._shortcut = "tem.Vacuum"

    @property
    def status(self) -> str:
        """ Status of the vacuum system. """
        return VacuumStatus(self._client.get(self._shortcut + ".Status")).name

    @property
    def is_buffer_running(self) -> bool:
        """ Checks whether the prevacuum pump is currently running
        (consequences: vibrations, exposure function blocked
        or should not be called).
        """
        return bool(self._client.get(self._shortcut + ".PVPRunning"))

    @property
    def is_column_open(self) -> bool:
        """ The status of the column valves. """
        return bool(self._client.get(self._shortcut + ".ColumnValvesOpen"))

    @property
    def gauges(self) -> Dict:
        """ Returns a dict with vacuum gauges information.
        Pressure values are in Pascals.
        """
        return self._client.call(self._shortcut + ".Gauges",
                                 obj=GaugesObj, func="show")

    def column_open(self) -> None:
        """ Open column valves. """
        self._client.set(self._shortcut + ".ColumnValvesOpen", True)

    def column_close(self) -> None:
        """ Close column valves. """
        self._client.set(self._shortcut + ".ColumnValvesOpen", False)

    def run_buffer_cycle(self) -> None:
        """ Runs a pumping cycle to empty the buffer. """
        self._client.call(self._shortcut + ".RunBufferCycle()")
