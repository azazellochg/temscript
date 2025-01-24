from ..utils.enums import VacuumStatus, GaugeStatus, GaugePressureLevel


class Vacuum:
    """ Vacuum functions. """
    def __init__(self, client):
        self._client = client

    @property
    def status(self) -> str:
        """ Status of the vacuum system. """
        return VacuumStatus(self._client.get("tem.Vacuum.Status")).name

    @property
    def is_buffer_running(self) -> bool:
        """ Checks whether the prevacuum pump is currently running
        (consequences: vibrations, exposure function blocked
        or should not be called).
        """
        return self._client.get("tem.Vacuum.PVPRunning")

    @property
    def is_column_open(self) -> bool:
        """ The status of the column valves. """
        return self._client.get("tem.Vacuum.ColumnValvesOpen")

    @property
    def gauges(self) -> dict:
        """ Returns a dict with vacuum gauges information.
        Pressure values are in Pascals.
        """
        gauges = {}
        for g in self._client.get("tem.Vacuum.Gauges"):
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

    def column_open(self) -> None:
        """ Open column valves. """
        self._client.set("tem.Vacuum.ColumnValvesOpen", True)

    def column_close(self) -> None:
        """ Close column valves. """
        self._client.set("tem.Vacuum.ColumnValvesOpen", False)

    def run_buffer_cycle(self) -> None:
        """ Runs a pumping cycle to empty the buffer. """
        self._client.call("tem.Vacuum.RunBufferCycle()")
