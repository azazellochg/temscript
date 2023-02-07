import socket
import json
from http.client import HTTPConnection
import logging

from .utils.enums import *
from .utils.marshall import ExtendedJsonEncoder, unpack_array, gzip_decode, MIME_TYPE_JSON
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
    def __init__(self, host="localhost", port=8001, timeout=None):
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
        #hasTemAdv = self._request("GET", "/v1/_tem_adv")[1]
        #useLD = self._request("GET", "/v1/_lowdose")[1]
        #useTecnaiCCD = self._request("GET", "/v1/_tecnai_ccd")[1]
        #useSEMCCD = self._request("GET", "/v1/_sem_ccd")[1]

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
