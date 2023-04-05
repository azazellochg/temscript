#!/usr/bin/python
import argparse
import platform
import logging
from http.server import BaseHTTPRequestHandler, HTTPServer

from .misc import *


class MicroscopeHandler(BaseHTTPRequestHandler):
    def get_microscope(self):
        """Return microscope object from server."""
        assert isinstance(self.server, MicroscopeServer)
        return self.server.microscope

    def build_response(self, response):
        """Encode response and send to client"""
        if response is None:
            self.send_response(204)
            self.end_headers()
            return

        try:
            encoded_response = ExtendedJsonEncoder().encode(response).encode("utf-8")
            # Compression?
            if len(encoded_response) > 256:
                encoded_response = gzip_encode(encoded_response)
                content_encoding = 'gzip'
            else:
                content_encoding = None
        except Exception as exc:
            logging.error("Exception raised during encoding of response: %s" % repr(exc))
            self.send_error(500, "Error handling request '%s': %s" % (self.path, str(exc)))
        else:
            self.send_response(200)
            if content_encoding:
                self.send_header('Content-Encoding', content_encoding)
            self.send_header('Content-Length', str(len(encoded_response)))
            self.send_header('Content-Type', MIME_TYPE_JSON)
            self.end_headers()
            self.wfile.write(encoded_response)

    def process_request(self, url, body=None):
        """ Get or set microscope attrs. """
        response = None
        microscope = self.get_microscope()

        # FIXME: change level to debug
        logging.info("RECV    url=%s" % url)
        if body is not None:
            logging.info("RECV params=%s" % body)

        if url.startswith("/get/"):
            response = rgetattr(microscope, url.split("/")[-1])
        elif url.startswith("/get_bool/"):
            result = rgetattr(microscope, url.split("/")[-1])
            response = True if result is not None else False
        elif url.startswith("/exec/"):
            response = rexecattr(microscope, url.split("/")[-1].rstrip("()"), body)
        elif url.startswith("/has/"):
            response = rhasattr(microscope, url.split("/")[-1])
        elif url.startswith("/set/"):
            url = url.split("/")[-1]
            value = body["value"]

            if "vector" in body:
                values = list(map(float, value))
                if len(values) != 2:
                    raise ValueError("Expected two values (X, Y) for Vector attribute %s" % url)

                if "limits" in body:
                    ranges = body["limits"]
                    for v in values:
                        if not (ranges[0] <= v <= ranges[1]):
                            raise ValueError("%s is outside of range %s" % (v, ranges))

                vector = rgetattr(microscope, url)
                vector.X = values[0]
                vector.Y = values[1]
                rsetattr(microscope, url, vector)
            else:
                rsetattr(microscope, url, value)

        else:
            raise ValueError("Invalid URL")

        # FIXME: change level to debug
        logging.info("REPLY      =%s" % response)

        return response

    def do_GET(self):
        """ Handler for the GET requests. """
        try:
            response = self.process_request(self.path)
        except AttributeError as exc:
            logging.error("AttributeError raised during handling of "
                          "GET request '%s': %s" % (self.path, repr(exc)))
            self.send_error(404, str(exc))
        except Exception as exc:
            logging.error("Exception raised during handling of "
                          "GET request '%s': %s" % (self.path, repr(exc)))
            self.send_error(500, "Error handling request '%s': %s" % (self.path, str(exc)))
        else:
            self.build_response(response)

    def do_POST(self):
        """ Handler for the POST requests. """
        try:
            length = int(self.headers['Content-Length'])
            if length > 4096:
                raise ValueError("Too much content...")
            content = self.rfile.read(length)
            decoded_content = json.loads(content.decode("utf-8"))
            response = self.process_request(self.path, decoded_content)
        except AttributeError as exc:
            logging.error("AttributeError raised during handling of "
                          "POST request '%s': %s" % (self.path, repr(exc)))
            self.send_error(404, str(exc))
        except Exception as exc:
            logging.error("Exception raised during  handling of "
                          "POST request '%s': %s" % (self.path, repr(exc)))
            self.send_error(500, "Error handling request '%s': %s" % (self.path, str(exc)))
        else:
            self.build_response(response)


class MicroscopeServer(HTTPServer, object):
    def __init__(self, server_address=('', 8080), useLD=True,
                 useTecnaiCCD=False, useSEMCCD=False):

        logging.basicConfig(level=logging.INFO,
                            datefmt='%d/%b/%Y %H:%M:%S',
                            format='[%(asctime)s] %(message)s',
                            handlers=[
                                logging.FileHandler("server.log", "w", "utf-8"),
                                logging.StreamHandler()])

        from ..base_microscope import BaseMicroscope
        self.microscope = BaseMicroscope(useLD, useTecnaiCCD, useSEMCCD)
        super().__init__(server_address, MicroscopeHandler)


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="This server should be started on the microscope PC",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument("-p", "--port", type=int, default=8080,
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
    parser.add_argument("--useSEMCCD", dest="useSEMCCD",
                        default=False, action='store_true',
                        help="Connect to SerialEMCCD plugin on Gatan PC that controls "
                             "Digital Micrograph (may be faster than via TIA / std scripting)")
    args = parser.parse_args(argv)

    if platform.system() != "Windows":
        raise NotImplementedError("This server should be started on the microscope PC (Windows only)")

    # Create a web server and define the handler to manage the incoming request
    server = MicroscopeServer((args.host, args.port),
                              useLD=args.useLD,
                              useTecnaiCCD=args.useTecnaiCCD,
                              useSEMCCD=args.useSEMCCD)
    try:
        logging.info("Started httpserver on host '%s' port %d." % (args.host, args.port))
        logging.info("Press Ctrl+C to stop server.")

        # Wait forever for incoming http requests
        server.serve_forever()

    except KeyboardInterrupt:
        logging.info("Ctrl+C received, shutting down the http server")

    finally:
        server.socket.close()

    return 0


if __name__ == '__main__':
    main()
