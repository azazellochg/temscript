import threading

from pytemscript.microscope import Microscope
from pytemscript.server.run import main as server_run


def run_server(argv):
    stop_event = threading.Event()
    thread = threading.Thread(target=server_run, args=(argv, stop_event))
    thread.start()
    return thread, stop_event

def test_interface(microscope):
    stage = microscope.stage
    print(stage.position)
    stage.go_to(x=1, y=-1)

def test_connection(connection_type="socket"):
    print("Testing %s connection" % connection_type)
    # Start server
    if connection_type == "socket":
        port = 39000
    elif connection_type == "grpc":
        port = 50051
    elif connection_type == "zmq":
        port = 5555
    else:
        raise Exception("Unknown connection type")

    args = ["-t", connection_type, "-p", port, "--host", "", "-d"]
    thread, stop_event = run_server(args)

    # Start client
    microscope = Microscope(connection=connection_type, host="",
                            port=port, debug=True)
    if microscope is None:
        raise RuntimeError("Could not create microscope client")

    test_interface(microscope)

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
