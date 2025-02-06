from .extras import Buttons


class UserButtons:
    """ User buttons control. """
    valid_buttons = ["L1", "L2", "L3", "R1", "R2", "R3"]

    def __init__(self, client):
        self._client = client

    @property
    def show(self) -> Buttons:
        """ Returns a dict with assigned hand panels buttons. """
        return self._client.call("tem.UserButtons", obj=Buttons, func="show")

    #TODO: add events - buttons assignment
