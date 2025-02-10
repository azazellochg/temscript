import pickle
try:
    import zmq
except ImportError:
    raise ImportError("Missing dependency 'pyzmq', please install it via pip")


class ZMQClient:
    def __init__(self, host, port):
        self.context = zmq.Context()
        self.socket = self.context.socket(zmq.REQ)  # Request-Reply pattern
        self.socket.connect(f"tcp://{host}:{port}")

    def call_method(self, method_name, *args, **kwargs):
        # Serialize the request
        message = pickle.dumps({
            'method': method_name,
            'args': args,
            'kwargs': kwargs
        })
        self.socket.send(message)
        # Receive the response
        response = self.socket.recv()
        return pickle.loads(response)
