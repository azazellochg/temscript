import argparse
import platform


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="This server should be started on the microscope PC",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument("-t", "--type", type=str,
                        choices=["socket", "zmq", "grpc"],
                        default="socket",
                        help="Server type to use: socket, zmq or grpc")
    parser.add_argument("-p", "--port", type=int,
                        default=39000,
                        help="Specify port on which the server is listening")
    parser.add_argument("--host", type=str, default='',
                        help="Specify host address on which the server is listening")
    parser.add_argument("--useLD", dest="useLD",
                        default=True, action='store_true',
                        help="Connect to LowDose server on microscope PC (limited control only)")
    parser.add_argument("--useTecnaiCCD", dest="useTecnaiCCD",
                        default=False, action='store_true',
                        help="Connect to TecnaiCCD plugin on microscope PC that controls "
                             "Digital Micrograph (may be faster than via TIA / std scripting)")
    args = parser.parse_args(argv)

    if platform.system() != "Windows":
        raise NotImplementedError("This server should be started on the microscope PC (Windows only)")

    if args.type == 'grpc':
        from .grpc_server import serve
        serve(args)
    elif args.type == 'zmq':
        from .zmq_server import ZMQServer
        server = ZMQServer(args)
        server.start()
    elif args.type == 'socket':
        from .socket_server import SocketServer
        server = SocketServer(args)
        server.start()


if __name__ == '__main__':
    main()
