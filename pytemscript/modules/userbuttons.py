class UserButtons:
    """ User buttons control. """
    valid_buttons = ["L1", "L2", "L3", "R1", "R2", "R3"]

    def __init__(self, client):
        self._client = client

    @property
    def list(self) -> dict:
        """ Returns a dict with assigned hand panels buttons. """
        buttons = self._client.get("tem.UserButtons")
        return {b.Name: b.Label for b in buttons}

    def __getattr__(self, name):
        if name in self.valid_buttons:
            button_index = self.valid_buttons.index(name)
            return self._client.get("tem.UserButtons")[button_index]
        raise AttributeError("Invalid button: %s" % name)
