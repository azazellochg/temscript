from typing import Dict
import math
import time
import logging

from ..utils.enums import MeasurementUnitType, StageStatus, StageHolderType, StageAxes
from .extras import StagePosition


class Stage:
    """ Stage functions. """
    def __init__(self, client):
        self._client = client
        self._err_msg = "Timeout. Stage is not ready"
        self._limits = dict()

    @property
    def _beta_available(self) -> bool:
        return self.limits['b']['unit'] != MeasurementUnitType.UNKNOWN.name

    def _wait_for_stage(self, tries: int = 10) -> None:
        """ Wait for stage to become ready. """
        attempt = 0
        while attempt < tries:
            if self._client.get("tem.Stage.Status") != StageStatus.READY:
                logging.info("Stage is not ready, waiting..")
                tries += 1
                time.sleep(1)
            else:
                break
        else:
            raise RuntimeError(self._err_msg)

    def _change_position(self,
                         direct: bool = False,
                         relative: bool = False,
                         **kwargs) -> None:
        """
        Execute stage move to a new position.

        :param direct: use Goto instead of MoveTo
        :param relative: use relative coordinates
        :param kwargs: new coordinates
        """
        self._wait_for_stage(tries=5)

        if relative:
            current_pos = self.position
            for axis in kwargs:
                kwargs[axis] += current_pos[axis]

        # convert units to meters and radians
        new_coords = dict()
        for axis in 'xyz':
            if kwargs.get(axis) is not None:

                new_coords.update({axis: kwargs[axis] * 1e-6})
        for axis in 'ab':
            if kwargs.get(axis) is not None:
                new_coords.update({axis: math.radians(kwargs[axis])})

        speed = kwargs.get("speed")
        if speed is not None and not (0.0 <= speed <= 1.0):
            raise ValueError("Speed must be within 0.0-1.0 range")

        if 'b' in new_coords and not self._beta_available:
            raise KeyError("B-axis is not available")

        limits = self.limits
        axes = 0
        for key, value in new_coords.items():
            if key not in 'xyzab':
                raise ValueError("Unexpected axis: %s" % key)
            if value < limits[key]['min'] or value > limits[key]['max']:
                raise ValueError('Stage position %s=%s is out of range' % (value, key))
            axes |= getattr(StageAxes, key.upper())

        # X and Y - 1000 to + 1000(micrometers)
        # Z - 375 to 375(micrometers)
        # a - 80 to + 80(degrees)
        # b - 29.7 to + 29.7(degrees)

        if not direct:
            self._client.call("tem.Stage", obj=StagePosition,
                              func="set", axes=axes,
                              method="MoveTo", **new_coords)
        else:
            if speed is not None:
                self._client.call("tem.Stage",
                                  obj=StagePosition, func="set",
                                  axes=axes, speed=speed,
                                  method="GoToWithSpeed", **new_coords)
            else:
                self._client.call("tem.Stage", obj=StagePosition,
                                  func="set", axes=axes,
                                  method="GoTo", **new_coords)

        self._wait_for_stage(tries=10)

    @property
    def status(self) -> str:
        """ The current state of the stage. """
        return StageStatus(self._client.get("tem.Stage.Status")).name

    @property
    def holder(self) -> str:
        """ The current specimen holder type. """
        return StageHolderType(self._client.get("tem.Stage.Holder")).name

    @property
    def position(self) -> Dict:
        """ The current position of the stage (x,y,z in um and a,b in degrees). """
        b = True if self._beta_available else False
        return self._client.call("tem.Stage.Position", obj=StagePosition,
                                   func="get", a=True, b=b)

    def go_to(self, relative=False, **kwargs) -> None:
        """ Makes the holder directly go to the new position by moving all axes
        simultaneously. Keyword args can be x,y,z,a or b.
        (x,y,z in um and a,b in degrees)

        :param relative: Use relative move instead of absolute position.
        :keyword float speed: fraction of the standard speed setting (max 1.0)
        """
        self._change_position(direct=True, relative=relative, **kwargs)

    def move_to(self, relative=False, **kwargs) -> None:
        """ Makes the holder safely move to the new position.
        Keyword args can be x,y,z,a or b.
        (x,y,z in um and a,b in degrees)

        :param relative: Use relative move instead of absolute position.
        """
        kwargs['speed'] = None
        self._change_position(relative=relative, **kwargs)

    @property
    def limits(self) -> Dict:
        """ Returns a dict with stage move limits. """
        if not self._limits:
            return self._client.call("tem.Stage", obj=StagePosition,
                                     func="limits")
        else:
            return self._limits
