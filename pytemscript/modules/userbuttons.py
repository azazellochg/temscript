import logging


class UserButtons:
    """ User buttons control. """
    valid_buttons = ["L1", "L2", "L3", "R1", "R2", "R3"]

    def __init__(self, client):
        self._client = client

    @property
    def list(self):
        """ Returns a dict with assigned hand panels buttons. """
        buttons = self._client.get("tem.UserButtons")
        return {b.Name: b.Label for b in buttons}

    def _get_button(self, name):
        button_index = self.valid_buttons.index(name)
        return self._client.get("tem.UserButtons")[button_index]

    def _assign_event(self, name, label, event_handler):
        button = self._get_button(name)

        def Pressed(name):
            logging.info("User button pressed: %s" % name)
            event_handler()

        button.Assignment = label

    def _remove_event(self, name):
        button = self._get_button(name)
        button.Assignment = ""

    def __getattr__(self, name):
        if name in self.valid_buttons:
            return self._get_button(name)
        raise AttributeError("Invalid button: %s" % name)

    def __setattr__(self, name, value):
        if name in self.valid_buttons:
            if isinstance(value, dict) and "label" in value and "method" in value:
                self._assign_event(name, value["label"], value["method"])
            else:
                raise ValueError("Value must be a dictionary with 'label' and 'method' keys.")
        else:
            super().__setattr__(name, value)

    def __delattr__(self, name):
        if name in self.valid_buttons:
            self._remove_event(name)
        else:
            super().__delattr__(name)
