import logging
import socket
import pickle
from typing import Dict, Any


class SocketClient:
    """ Remote socket client interface for the microscope.

    :param host: Remote hostname or IP address
    :type host: str
    :param port: Remote port number
    :type port: int
    :param debug: Print debug messages
    :type debug: bool
    """
    def __init__(self,
                 host: str = "localhost",
                 port: int = 39000,
                 debug: bool = False):
        self.host = host
        self.port = port

        logging.basicConfig(level=logging.DEBUG if debug else logging.INFO,
                            datefmt='%d/%b/%Y %H:%M:%S',
                            format='[CLIENT] [%(asctime)s] %(message)s',
                            handlers=[
                                logging.FileHandler("socket_client.log", "w", "utf-8"),
                                logging.StreamHandler()])

    def __getattr__(self, method_name: str):
        """ This handles both method calls and properties of the Microscope client instance. """
        def method_or_property(*args, **kwargs):
            payload = {
                "method": method_name,
                "args": args,
                "kwargs": kwargs
            }
            logging.debug("Sending request: %s" % payload)
            response = self.__send_request(payload)
            logging.debug("Received response: %s" % response)
            return response

        return method_or_property

    def __send_request(self, payload: Dict[str, Any]):
        """ Send data to the remote server and return response. """
        try:
            with socket.create_connection((self.host, self.port)) as client_socket:
                serialized_data = pickle.dumps(payload)
                client_socket.sendall(len(serialized_data).to_bytes(4, byteorder="big") + serialized_data)

                response_length_bytes = client_socket.recv(4)
                response_length = int.from_bytes(response_length_bytes, byteorder='big')

                response_data = client_socket.recv(response_length)
                if not response_data:
                    raise ConnectionError("No response received from server")

                return pickle.loads(response_data)

        except Exception as e:
            raise RuntimeError("Error communicating with server: %s" % e)
