import threading
import time
from typing import List

from pytemscript.microscope import Microscope
from pytemscript.server.run import main as server_run


def server_thread(argv: List[str]):
    """ Start server process in a separate thread. """
    stop_event = threading.Event()
    thread = threading.Thread(target=server_run, args=(argv, stop_event))
    thread.start()
    return thread, stop_event

def test_interface(microscope):
    """ Test remote interface. """
    stage = microscope.stage
    print(stage.position)
    stage.go_to(x=1, y=-1)

def test_connection(connection_type: str = "socket"):
    """ Create server and client, then test the connection. """
    print("Testing %s connection" % connection_type)
    if connection_type == "socket":
        port = 39000
    elif connection_type == "grpc":
        port = 50051
    elif connection_type == "zmq":
        port = 5555
    else:
        raise ValueError("Unknown connection type")

    # Start server
    args = ["-t", connection_type, "-p", port, "--host", "127.0.0.1", "-d"]
    thread, stop_event = server_thread(args)
    time.sleep(1)

    # Start client
    client = Microscope(connection=connection_type, host="",
                        port=port, debug=True)
    if client is None:
        raise RuntimeError("Could not create microscope client")

    test_interface(client)

    # Stop server
    stop_event.set()
    thread.join()

def main():
    """ Basic test to check server-client connection on localhost. """
    test_connection(connection_type="socket")
    test_connection(connection_type="grpc")
    test_connection(connection_type="zmq")

if __name__ == '__main__':
    main()
