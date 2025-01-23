from .modules import *
from .utils.enums import ProductFamily, CondenserLensSystem


class Microscope:
    """ Base client interface, exposing available methods
     and properties.
    """
    def __init__(self, connection="direct", *args, **kwargs):
        self._communication_type = connection
        if connection == "direct":
            from .clients.com_client import COMClient
            self.client = COMClient(*args, **kwargs)
        elif connection == 'grpc':
            self.client = GRPCClient(*args, **kwargs)
        elif connection == 'zmq':
            self.client = ZMQClient(*args, **kwargs)
        elif connection == 'socket':
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
        self.user_buttons = UserButtons(client)

        if client.has_advanced_iface:
            self.user_door = UserDoor(client)
            self.energy_filter = EnergyFilter(client)

        if kwargs.get("useLD", False):
            self.low_dose = LowDose(client)

    @property
    def family(self):
        """ Returns the microscope product family / platform. """
        value = self.client.get_from_cache("tem.Configuration.ProductFamily")
        return ProductFamily(value).name

    @property
    def condenser_system(self):
        """ Returns the type of condenser lens system: two or three lenses. """
        value = self.client.get_from_cache("tem.Configuration.CondenserLensSystem")
        return CondenserLensSystem(value).name
