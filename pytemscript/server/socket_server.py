from argparse import Namespace
import socket
import pickle
import logging


class SocketServer:
    """ Simple UNIX socket server. Not secure at all. """
    def __init__(self, args: Namespace):
        """ Initialize the basic variables and logging. """
        self.server_socket = None
        self.server_com = None
        self.host = args.host or "localhost"
        self.port = args.port or 39000
        self.useLD = args.useLD
        self.useTecnaiCCD = args.useTecnaiCCD

        logging.basicConfig(level=logging.DEBUG if args.debug else logging.INFO,
                            datefmt='%d/%b/%Y %H:%M:%S',
                            format='[SERVER] [%(asctime)s] %(message)s',
                            handlers=[
                                logging.FileHandler("socket_server.log", "w", "utf-8"),
                                logging.StreamHandler()])

    def start(self):
        """ Start both the COM client (as a server) and the socket server. """
        try:
            from ..clients.com_client import COMClient
            self.server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.server_socket.bind((self.host, self.port))
            self.server_socket.listen(5)
            logging.info("Socket server listening on %s:%d" % (self.host, self.port))

            while True:
                try:
                    # start COM client as a server
                    self.server_com = COMClient(useTecnaiCCD = self.useTecnaiCCD,
                                                useLD=self.useLD,
                                                as_server=True)
                    client_socket, client_address = self.server_socket.accept()
                    logging.info("Connection received from: " + str(client_address))
                    with client_socket:
                        data = client_socket.recv(4096)
                        message = pickle.loads(data)
                        method_name = message['method']
                        args = message['args']
                        kwargs = message['kwargs']
                        # Call the appropriate method and send back the result
                        logging.debug("Received request: %s, args: %s, kwargs: %s" % (
                            method_name, args, kwargs))
                        result = self.handle_request(method_name, *args, **kwargs)
                        logging.debug("Sending response: %s" % result)
                        client_socket.send(pickle.dumps(result))
                except socket.error as e:
                    logging.error(e)

        except KeyboardInterrupt:
            logging.info("Ctrl+C received. Server shutting down..")
        finally:
            if self.server_socket:
                self.server_socket.close()
                # explicitly stop COM server
                self.server_com._scope._close()
                self.server_com = None

    def handle_request(self, method_name, *args, **kwargs):
        """ Process a socket message: pass method to the COM server
         and return result to the client. """
        method = getattr(self.server_com, method_name, None)
        if method is None:
            raise ValueError("Unknown method: %s" % method_name)
        elif callable(method):
            return method(*args, **kwargs)
        else:  # for property decorators
            return method
