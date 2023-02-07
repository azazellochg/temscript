import socket
import json
from http.client import HTTPConnection
from urllib.parse import urlencode, quote_plus
import logging

from .utils.enums import *
from .utils.marshall import ExtendedJsonEncoder, unpack_array, gzip_decode, MIME_TYPE_JSON


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
        self.port = port
        self.host = host
        self.timeout = timeout
        self._conn = None

        logging.basicConfig(level=logging.INFO,
                            datefmt='%d-%m-%Y %H:%M:%S',
                            format='%(asctime)s %(message)s',
                            handlers=[
                                logging.FileHandler("debug.log", "w", "utf-8"),
                                logging.StreamHandler()])

        # Make connection
        self._conn = HTTPConnection(self.host, self.port, timeout=self.timeout)

    def _request(self, method, endpoint, body=None, query=None, headers=None, accepted_response=None):
        """
        Send request to server.

        If accepted_response is None, 200 is accepted for all methods, additionally 204 is accepted
        for methods, which pass a body (PUT, PATCH, POST)

        :param method: HTTP method to use, e.g. "GET" or "PUT"
        :type method: str
        :param endpoint: URL to request
        :type endpoint: str
        :param query: Optional dict or iterable of key-value-tuples to encode as query
        :type: Union[Dict[str, str], Iterable[Tuple[str, str]], None]
        :param body: Body to send
        :type body: Optional[Union[str, bytes]]
        :param headers: Optional dict of additional headers.
        :type headers: Optional[Dict[str, str]]
        :param accepted_response: Accepted response codes
        :type accepted_response: Optional[List[int]]

        :returns: response, decoded response body
        """
        if accepted_response is None:
            accepted_response = [200]
            if method in ["PUT", "PATCH", "POST"]:
                accepted_response.append(204)

        # Create request
        headers = dict(headers) if headers is not None else dict()
        if "Accept" not in headers:
            headers["Accept"] = MIME_TYPE_JSON
        if "Accept-Encoding" not in headers:
            headers["Accept-Encoding"] = "gzip"
        if query is not None:
            url = endpoint + '?' + urlencode(query)
        else:
            url = endpoint
        self._conn.request(method, url, body, headers)

        # Get response
        try:
            response = self._conn.getresponse()
        except socket.timeout:
            self._conn.close()
            self._conn = None
            raise

        if response.status not in accepted_response:
            if response.status == 404:
                raise KeyError("Failed remote call: %s" % response.reason)
            else:
                raise RuntimeError("Remote call returned status %d: %s" % (response.status, response.reason))
        if response.status == 204:
            return response, None

        # Decode response
        content_length = response.getheader("Content-Length")
        if content_length is not None:
            encoded_body = response.read(int(content_length))
        else:
            encoded_body = response.read()

        content_type = response.getheader("Content-Type")
        if content_type != MIME_TYPE_JSON:
            raise ValueError("Unexpected response type: %s", content_type)
        if response.getheader("Content-Encoding") == "gzip":
            encoded_body = gzip_decode(encoded_body)

        body = json.loads(encoded_body.decode("utf-8"))
        return response, body

    def _request_with_json_body(self, method, url, body, query=None, headers=None, accepted_response=None):
        """
        Like :meth:`_request` but body is encoded as JSON.

        ..see :: :meth:`_request`
        """
        if body is not None:
            headers = dict(headers) if headers is not None else dict()
            headers["Content-Type"] = MIME_TYPE_JSON
            encoder = ExtendedJsonEncoder()
            encoded_body = encoder.encode(body).encode("utf-8")
        else:
            encoded_body = None
        return self._request(method, url, body=encoded_body, query=query, headers=headers,
                             accepted_response=accepted_response)

    def get_family(self):
        return self._request("GET", "/v1/family")[1]

    def set_column_valves_open(self, state):
        self._request_with_json_body("PUT", "/v1/column_valves_open", state)
