from typing import Dict

from .extras import SpecialObj


class ButtonsObj(SpecialObj):
    """ Wrapper around buttons COM object. """

    def show(self) -> Dict:
        """ Returns a dict with buttons assignment. """
        buttons = {}
        for b in self.com_object:
            buttons[b.Name] = b.Label

        return buttons


class UserButtons:
    """ User buttons control. """
    valid_buttons = ["L1", "L2", "L3", "R1", "R2", "R3"]

    def __init__(self, client):
        self._client = client

    @property
    def show(self) -> Dict:
        """ Returns a dict with assigned hand panels buttons. """
        return self._client.call("tem.UserButtons", obj=ButtonsObj, func="show")

    #TODO: add events - buttons assignment
