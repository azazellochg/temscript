from modules import *
from utils.enums import ProductFamily, CondenserLensSystem


class Microscope:
    """ Base client interface, exposing available methods
     and properties.
    """
    def __init__(self, communication_type="direct", *args, **kwargs):
        self._communication_type = communication_type
        if communication_type == "direct":
            from clients.com_client import COMClient
            self.client = COMClient(*args, **kwargs)
        elif communication_type == 'grpc':
            self.client = GRPCClient(*args, **kwargs)
        elif communication_type == 'zmq':
            self.client = ZMQClient(*args, **kwargs)
        elif communication_type == 'pure_python':
            self.client = SocketClient(*args, **kwargs)
        else:
            raise ValueError("Unsupported communication type")

        client = self.client
        self._cache = self.client.cache

        self.acquisition = Acquisition(client)
        self.detectors = Detectors(client)
        self.gun = Gun(client)
        self.optics = Optics(client, self.condenser_system)
        self.stem = Stem(client)
        self.vacuum = Vacuum(client)
        self.autoloader = Autoloader(client)
        self.stage = Stage(client)
        self.piezo_stage = PiezoStage(client)
        self.apertures = Apertures(client)
        self.temperature = Temperature(client)

        if client.has_advanced_iface:
            self.user_door = UserDoor(client)
            self.energy_filter = EnergyFilter(client)

        if kwargs.get("useLD", False):
            self.low_dose = LowDose(client)

    def _check_cache(self, key, fetch_func):
        if key not in self._cache:
            self._cache[key] = fetch_func()
        return self._cache[key]

    def get(self, attr):
        return self.client.get(attr)

    def set(self, attr, value, **kwargs):
        return self.client.set(attr, value, **kwargs)

    def has(self, attr):
        """ Get request with cache support. Should be used only for attributes
        that do not change over the session. """
        if attr not in self._cache:
            try:
                _ = self.get(attr)
                self._cache[attr] = True
            except AttributeError:
                self._cache[attr] = False

        return self._cache.get(attr)

    def call(self, attr, *args, **kwargs):
        return self.client.call(attr, *args, **kwargs)

    @property
    def family(self):
        """ Returns the microscope product family / platform. """
        return self._check_cache("family",
                                 lambda: ProductFamily(self.get("tem.Configuration.ProductFamily")).name)

    @property
    def condenser_system(self):
        """ Returns the type of condenser lens system: two or three lenses. """
        return self._check_cache("condenser_system",
                                 lambda: CondenserLensSystem(self.get("tem.Configuration.CondenserLensSystem")).name)

    @property
    def user_buttons(self):
        """ Returns a dict with assigned hand panels buttons. """
        return self._check_cache("user_buttons",
                                    lambda: {b.Name: b.Label for b in self.get("tem.UserButtons")})
