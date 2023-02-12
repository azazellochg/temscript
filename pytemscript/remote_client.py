import socket
import json
from http.client import HTTPConnection

from .objects import *
from .utils.enums import *
from .utils.misc import ExtendedJsonEncoder, unpack_array, gzip_decode, MIME_TYPE_JSON


class RemoteMicroscope:
    """ Remote client interface to the microscope.

    :param host: Specify host address on which the server is listening
    :type host: IP address or hostname
    :param port: Specify port on which the server is listening
    :type port: int
    :param timeout: Connection timeout in seconds
    :type timeout: int
    """
    def __init__(self, host="localhost", port=8080, timeout=None):
        self._host = host
        self._port = port
        self._timeout = timeout
        self._conn = None

        self._conn = HTTPConnection(self._host, self._port, timeout=self._timeout)

        self._has_tem_adv = self.has("tem_adv")
        self._useLD = self.has("tem_lowdose")
        #self._useTecnaiCCD = self.has("_tecnai_ccd")
        #self._useSEMCCD = self.has("_sem_ccd")

        #self.acquisition = Acquisition(self)
        #self.detectors = Detectors(self)
        self.gun = Gun(self)
        self.optics = Optics(self)
        self.stem = Stem(self)
        self.temperature = Temperature(self)
        self.vacuum = Vacuum(self)
        self.autoloader = Autoloader(self)
        self.stage = Stage(self)
        self.piezo_stage = PiezoStage(self)
        #self.apertures = Apertures(self)

        if self._has_tem_adv:
            self.user_door = UserDoor(self)
            self.energy_filter = EnergyFilter(self)

        if self._useLD:
            self.lowdose = LowDose(self)

    def _request(self, method, endpoint, body=None):
        """
        Send request to server.

        :param method: HTTP method to use, e.g. "GET" or "POST"
        :type method: str
        :param endpoint: URL to request
        :type endpoint: str
        :param body: Body to send
        :type body: Optional[Union[str, bytes]]
        :returns: response, decoded response body
        """
        # Create request
        if body is not None:
            encoder = ExtendedJsonEncoder()
            body = encoder.encode(body).encode("utf-8")

        headers = {
            "Accept": MIME_TYPE_JSON,
            "Accept-Encoding": "gzip"
        }

        self._conn.request(method, endpoint, body, headers)

        # Get response
        try:
            response = self._conn.getresponse()
        except socket.timeout:
            self._conn.close()
            self._conn = None
            raise

        if response.status == 204:  # returns nothing, e.g. after a successfull SET
            return response, None

        if response.status != 200:
            raise RuntimeError("Failed remote call: %s" % response.reason)

        # Decode response
        content_length = response.getheader("Content-Length")
        if content_length is not None:
            encoded_body = response.read(int(content_length))
        else:
            encoded_body = response.read()

        if response.getheader("Content-Encoding") == "gzip":
            encoded_body = gzip_decode(encoded_body)

        body = json.loads(encoded_body.decode("utf-8"))

        return body

    def get(self, attrname):
        return self._request("GET", "/get/" + attrname)

    def exec(self, attrname, *args, **kwargs):
        if args is not None:
            return self._request("POST", "/exec/" + attrname, args)
        else:
            return self._request("POST", "/exec/" + attrname, kwargs)

    def has(self, attrname):
        return self._request("GET", "/has/" + attrname)

    def set(self, attrname, value, vector=False, limits=None):
        body = {"value": value}
        if vector:
            body["vector"] = True
            if limits is not None:
                body["limits"] = limits

        self._request("POST", "/set/" + attrname, body)

    @property
    def family(self):
        """ Returns the microscope product family / platform. """
        return ProductFamily(self.get("tem.Configuration.ProductFamily")).name

    @property
    def condenser_system(self):
        """ Returns the type of condenser lens system: two or three lenses. """
        return CondenserLensSystem(self.get("tem.Configuration.CondenserLensSystem")).name

    @property
    def user_buttons(self):
        """ Returns a dict with assigned hand panels buttons. """
        return {b.Name: b.Label for b in self.get("tem.UserButtons")}
