#!/usr/bin/python
import json
import functools
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import urlparse
import argparse
import platform

from .marshall import ExtendedJsonEncoder, gzip_encode, MIME_TYPE_JSON, pack_array


def multi_getattr(obj, attr):
    attributes = attr.split(".")
    for i in attributes:
        try:
            obj = getattr(obj, i)
            if callable(obj):
                obj = obj()
        except AttributeError:
            raise
    return obj


def rsetattr(obj, attr, val):
    """ https://stackoverflow.com/a/31174427 """
    pre, _, post = attr.rpartition('.')
    return setattr(rgetattr(obj, pre) if pre else obj, post, val)


def rgetattr(obj, attr, kwargs=None, is_callable=False):
    def _getattr(obj, attr):
        return getattr(obj, attr)
    result = functools.reduce(_getattr, [obj] + attr.split('.'))
    if is_callable:
        if kwargs is not None:
            return result(kwargs)
        else:
            return result()
    else:
        return result


def rhasattr(obj, attr):
    """ https://stackoverflow.com/a/65781864 """
    try:
        functools.reduce(getattr, attr.split("."), obj)
        return True
    except AttributeError:
        return False


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
            self.log_error("Exception raised during encoding of response: %s" % repr(exc))
            self.send_error(500, "Error handling request '%s': %s" % (self.path, str(exc)))
        else:
            self.send_response(200)
            if content_encoding:
                self.send_header('Content-Encoding', content_encoding)
            self.send_header('Content-Length', str(len(encoded_response)))
            self.send_header('Content-Type', MIME_TYPE_JSON)
            self.end_headers()
            self.wfile.write(encoded_response)

    def process_request(self, request, body=None):
        """ Get or set microscope attrs. """
        response = None
        url = request.path
        microscope = self.get_microscope()

        print("url=", url)
        print("params=", body)

        if url.startswith("/get/"):
            response = rgetattr(microscope, url.lstrip("/get/"))
        elif url.startswith("/exec/"):
            response = rgetattr(microscope, url.lstrip("/exec/").rstrip("()"), body, is_callable=True)
        elif url.startswith("/set/"):
            rsetattr(microscope, url.lstrip("/set/"), body)
        elif url.startswith("/has/"):
            response = rhasattr(microscope, url.lstrip("/has/"))
        else:
            raise RuntimeError("URL %s not found" % url)

        return response

    def do_GET(self):
        """ Handler for the GET requests. """
        try:
            request = urlparse(self.path)
            response = self.process_request(request)
        except AttributeError as exc:
            self.log_error("AttributeError raised during handling of GET request '%s': %s" % (self.path, repr(exc)))
            self.send_error(404, str(exc))
        except Exception as exc:
            self.log_error("Exception raised during handling of GET request '%s': %s" % (self.path, repr(exc)))
            self.send_error(500, "Error handling request '%s': %s" % (self.path, str(exc)))
        else:
            self.build_response(response)

    def do_POST(self):
        """ Handler for the POST requests. """
        try:
            request = urlparse(self.path)
            length = int(self.headers['Content-Length'])
            if length > 4096:
                raise ValueError("Too much content...")
            content = self.rfile.read(length)
            decoded_content = json.loads(content.decode("utf-8"))
            response = self.process_request(request, decoded_content)
        except AttributeError as exc:
            self.log_error("AttributeError raised during handling of POST request '%s': %s" % (self.path, repr(exc)))
            self.send_error(404, str(exc))
        except Exception as exc:
            self.log_error("Exception raised during handling of POST request '%s': %s" % (self.path, repr(exc)))
            self.send_error(500, "Error handling request '%s': %s" % (self.path, str(exc)))
        else:
            self.build_response(response)


class MicroscopeServer(HTTPServer, object):
    def __init__(self, server_address=('', 8080), useLD=False, useTecnaiCCD=False, useSEMCCD=False):
        from pytemscript.microscope import Microscope
        self.microscope = Microscope(useLD, useTecnaiCCD, useSEMCCD)
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
        print("Started httpserver on host '%s' port %d." % (args.host, args.port))
        print("Press Ctrl+C to stop server.")

        # Wait forever for incoming http requests
        server.serve_forever()

    except KeyboardInterrupt:
        print('Ctrl+C received, shutting down the http server')

    finally:
        server.socket.close()

    return 0


if __name__ == '__main__':
    main()
