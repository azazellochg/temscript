from argparse import Namespace
import socket
import pickle


class SocketServer:
    def __init__(self, args: Namespace):
        self.host = args.host or "localhost"
        self.port = args.port or 39000
        self.useLD = args.useLD
        self.useTecnaiCCD = args.useTecnaiCCD

    def start(self):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as server_socket:
            server_socket.bind((self.host, self.port))
            server_socket.listen(5)
            while True:
                client_socket, client_address = server_socket.accept()
                with client_socket:
                    data = client_socket.recv(4096)
                    message = pickle.loads(data)
                    method_name = message['method']
                    args = message['args']
                    kwargs = message['kwargs']
                    # Call the appropriate method and send back the result
                    result = call_method_on_server_object(method_name, *args, **kwargs)
                    client_socket.send(pickle.dumps(result))
