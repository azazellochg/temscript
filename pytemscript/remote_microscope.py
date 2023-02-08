import socket
import json
from http.client import HTTPConnection
import logging

from .utils.enums import *
from .utils.marshall import ExtendedJsonEncoder, unpack_array, gzip_decode, MIME_TYPE_JSON
#from .base_microscope import Vector
from .microscope import (Acquisition, Detectors, Gun, LowDose, Optics, Stem, Temperature,
                         Vacuum, Autoloader, Stage, PiezoStage, Apertures, UserDoor, EnergyFilter)


class RemoteMicroscope:
    """ High level interface to the remote microscope server.
    Boolean params below have to match the server params.
    :param host: Specify host address on which the server is listening
    :type host: IP address or hostname
    :param port: Specify port on which the server is listening
    :type port: int
    :param timeout: Connection timeout in seconds
    :type timeout: int
    """
    def __init__(self, host="localhost", port=8080, timeout=None):
        self._port = port
        self._host = host
        self._timeout = timeout
        self._conn = None

        logging.basicConfig(level=logging.INFO,
                            datefmt='%d-%m-%Y %H:%M:%S',
                            format='%(asctime)s %(message)s',
                            handlers=[
                                logging.FileHandler("remote_debug.log", "w", "utf-8"),
                                logging.StreamHandler()])

        # Make connection
        self._conn = HTTPConnection(self._host, self._port, timeout=self._timeout)

        hasTem = self._request("GET", "/has/_tem")[1]
        print("Received hasTem=", hasTem)
        #hasTemAdv = self._request("GET", "/has/_tem_adv")[1]
        #useLD = self._request("GET", "/has/_lowdose")[1]
        #useTecnaiCCD = self._request("GET", "/has/_tecnai_ccd")[1]
        #useSEMCCD = self._request("GET", "/has/_sem_ccd")[1]

        #self.user_door = UserDoor(self)

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

        return response, body

    @property
    def family(self):
        """ Returns the microscope product family / platform. """
        result = self._request("GET", "/get/_tem.Configuration.ProductFamily")[1]
        return ProductFamily(result).name

    @property
    def intensity(self):
        """ Intensity / C2 condenser lens value. (read/write)"""
        return self._request("GET", "/get/_tem.Illumination.Intensity")[1]

    @intensity.setter
    def intensity(self, value):
        if not (0.0 <= value <= 1.0):
            raise ValueError("%s is outside of range 0.0-1.0" % value)
        self._request("POST", "/set/_tem.Illumination.Intensity", float(value))

    @property
    def beam_shift(self):
        """ Beam shift X and Y in um. (read/write)"""
        x = float(self._request("GET", "/get/_tem.Illumination.Shift.X")[1]) * 1e6
        y = float(self._request("GET", "/get/_tem.Illumination.Shift.Y")[1]) * 1e6
        return (x, y)

    #@beam_shift.setter
    #def beam_shift(self, value):
    #    new_value = (value[0] * 1e-6, value[1] * 1e-6)
    #    Vector.set(self._tem_illumination, "Shift", new_value)

    def run_buffer_cycle(self):
        """ Runs a pumping cycle to empty the buffer. """
        self._request("GET", "/exec/_tem.Vacuum.RunBufferCycle()")

    def column_close(self):
        """ Close column valves. """
        self._request("POST", "/set/_tem.Vacuum.ColumnValvesOpen", True)

    def normalize(self, mode):
        """ Normalize condenser or projection lens system.
        :param mode: Normalization mode (ProjectionNormalization or IlluminationNormalization enum)
        :type mode: IntEnum
        """
        if mode in ProjectionNormalization:
            self._request("POST", "/exec/_tem.Projection.Normalize()", mode)
        elif mode in IlluminationNormalization:
            self._request("POST", "/exec/_tem.Illumination.Normalize()", mode)
        else:
            raise ValueError("Unknown normalization mode: %s" % mode)
