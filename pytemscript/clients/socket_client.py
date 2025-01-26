import logging
import socket
import pickle


class SocketClient:
    def __init__(self,
                 host: str = "localhost",
                 port: int = 39000,
                 debug: bool = False):
        self.host = host
        self.port = port

        logging.basicConfig(level=logging.DEBUG if debug else logging.INFO,
                            datefmt='%d/%b/%Y %H:%M:%S',
                            format='[%(asctime)s] %(message)s',
                            handlers=[
                                logging.FileHandler("socket_client.log", "w", "utf-8"),
                                logging.StreamHandler()])

    def call_method(self, method_name, *args, **kwargs):
        # Serialize the request
        message = pickle.dumps({
            'method': method_name,
            'args': args,
            'kwargs': kwargs
        })
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.connect((self.host, self.port))
            s.sendall(message)
            # Receive the response
            data = s.recv(4096)
            return pickle.loads(data)
