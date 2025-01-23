class UserButtons:
    """ User buttons control. """
    def __init__(self, client):
        self._client = client
        self._valid_buttons = ["L1", "L2", "L3", "R1", "R2", "R3"]
        self._event_handlers = {}

    @property
    def list(self):
        """ Returns a dict with assigned hand panels buttons. """
        buttons = self._client.get("tem.UserButtons")
        return {b.Name: b.Label for b in buttons}

    def _get_button(self, name):
        return self._client.get("tem.UserButtons." + name)

    def _assign_event(self, name, label, event_handler):
        button = self._get_button(name)

        def wrapper(*args, **kwargs):
            event_handler(*args, **kwargs)

        button.Assignment = label
        button.Pressed = wrapper
        self._event_handlers[name] = event_handler

    def _remove_event(self, name):
        button = self._get_button(name)
        button.Assignment = ""
        if name in self._event_handlers:
            del self._event_handlers[name]

    def __getattr__(self, name):
        if name in self._valid_buttons:
            return self._get_button(name)
        raise KeyError("Invalid button name: %s" % name)

    def __setattr__(self, name, value):
        if name in self._valid_buttons:
            if isinstance(value, dict) and "label" in value and "method" in value:
                self._assign_event(name, value["label"], value["method"])
            else:
                raise ValueError("Value must be a dictionary with 'label' and 'method' keys.")
        else:
            raise KeyError("Invalid button name: %s" % name)

    def __delattr__(self, name):
        if name in self._valid_buttons:
            self._remove_event(name)
        else:
            raise KeyError("Invalid button name: %s" % name)
