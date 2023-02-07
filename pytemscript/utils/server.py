#!/usr/bin/python
import json
import functools
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import urlparse, parse_qs, unquote
import argparse
import platform

from .marshall import ExtendedJsonEncoder, gzip_encode, MIME_TYPE_PICKLE, MIME_TYPE_JSON, pack_array


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


def rgetattr(obj, attr, *args):
    def _getattr(obj, attr):
        return getattr(obj, attr, *args)
    result = functools.reduce(_getattr, [obj] + attr.split('.'))
    if callable(result):
        result()
        return "OK"
    else:
        return result


def rhasattr(obj, attr):
    """ https://stackoverflow.com/a/65781864 """
    try:
        functools.reduce(getattr, attr.split("."), obj)
        return True
    except AttributeError:
        return False


class TestScope:
    def __init__(self):
        print("Initialised TestScope")
        self._tem_adv = True

    @property
    def family(self):
        return 1


class MicroscopeHandler(BaseHTTPRequestHandler):
    GET_V1_FORWARD = ("family", "microscope_id", "version", "voltage", "vacuum", "stage_holder",
                      "stage_status", "stage_position", "stage_limits", "detectors", "cameras", "stem_detectors",
                      "stem_acquisition_param", "image_shift", "beam_shift", "beam_tilt", "projection_sub_mode",
                      "projection_mode", "projection_mode_string", "magnification_index", "indicated_camera_length",
                      "indicated_magnification", "defocus", "objective_excitation", "intensity", "objective_stigmator",
                      "condenser_stigmator", "diffraction_shift", "screen_current", "screen_position",
                      "illumination_mode", "condenser_mode", "illuminated_area", "probe_defocus", "convergence_angle",
                      "stem_magnification", "stem_rotation", "spot_size_index", "dark_field_mode", "beam_blanked",
                      "instrument_mode", 'optics_state', 'state', 'column_valves_open')

    PUT_V1_FORWARD = ("image_shift", "beam_shift", "beam_tilt", "projection_mode", "magnification_index",
                      "defocus", "intensity", "diffraction_shift", "objective_stigmator", "condenser_stigmator",
                      "screen_position", "illumination_mode", "spot_size_index", "dark_field_mode",
                      "condenser_mode", "illuminated_area", "probe_defocus", "convergence_angle",
                      "stem_magnification", "stem_rotation", "beam_blanked", "instrument_mode")

    def get_microscope(self):
        """Return microscope object from server."""
        assert isinstance(self.server, MicroscopeServer)
        return self.server.microscope

    def get_accept_types(self):
        """Return list of accepted encodings."""
        return [x.split(';', 1)[0].strip() for x in self.headers.get("Accept", "").split(",")]

    def build_response(self, response):
        """Encode response and send to client"""
        if response is None:
            self.send_response(204)
            self.end_headers()
            return

        try:
            accept_type = self.get_accept_types()
            if MIME_TYPE_PICKLE in accept_type:
                import pickle
                encoded_response = pickle.dumps(response, protocol=2)
                content_type = MIME_TYPE_PICKLE
            else:
                encoded_response = ExtendedJsonEncoder().encode(response).encode("utf-8")
                content_type = MIME_TYPE_JSON

            # Compression?
            accept_encoding = [x.split(';', 1)[0].strip() for x in self.headers.get("Accept-Encoding", "").split(",")]
            if len(encoded_response) > 256 and 'gzip' in accept_encoding:
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
            self.send_header('Content-Type', content_type)
            self.end_headers()
            self.wfile.write(encoded_response)

    def do_GET_V1(self, endpoint, query):
        """Handle V1 GET requests"""
        if endpoint in self.GET_V1_FORWARD:
            response = getattr(self.get_microscope(), 'get_' + endpoint)()
        elif endpoint.startswith("detector_param/"):
            name = unquote(endpoint[15:])
            response = self.get_microscope().get_detector_param(name)
        elif endpoint.startswith("camera_param/"):
            name = unquote(endpoint[13:])
            response = self.get_microscope().get_camera_param(name)
        elif endpoint.startswith("stem_detector_param/"):
            name = unquote(endpoint[20:])
            response = self.get_microscope().get_stem_detector_param(name)
        elif endpoint == "acquire":
            detectors = tuple(query.get("detectors", ()))
            response = self.get_microscope().acquire(*detectors)
            if MIME_TYPE_PICKLE not in self.get_accept_types():
                response = {key: pack_array(value) for key, value in response.items()}
        elif endpoint == "stem_available":
            response = self.get_microscope().is_stem_available()
        else:
            raise KeyError("Unknown endpoint: '%s'" % endpoint)
        return response

    def do_PUT_V1(self, endpoint, query):
        """Handle V1 PUT requests"""
        length = int(self.headers['Content-Length'])
        if length > 4096:
            raise ValueError("Too much content...")
        content = self.rfile.read(length)
        decoded_content = json.loads(content.decode("utf-8"))

        # Check for known endpoints
        response = None
        if endpoint in self.PUT_V1_FORWARD:
            response = getattr(self.get_microscope(), 'set_' + endpoint)(decoded_content)
        elif endpoint == "stage_position":
            method = str(query["method"][0]) if "method" in query else None
            speed = float(query["speed"][0]) if "speed" in query else None
            pos = dict((k, decoded_content[k]) for k in decoded_content.keys() if k in STAGE_AXES)
            self.get_microscope().set_stage_position(pos, method=method, speed=speed)
        elif endpoint.startswith("camera_param/"):
            name = unquote(endpoint[13:])
            ignore_errors = bool(query.get("ignore_errors", [False])[0])
            response = self.get_microscope().set_camera_param(name, decoded_content, ignore_errors=ignore_errors)
        elif endpoint.startswith("stem_detector_param/"):
            name = unquote(endpoint[20:])
            ignore_errors = bool(query.get("ignore_errors", [False])[0])
            response = self.get_microscope().set_stem_detector_param(name, decoded_content, ignore_errors=ignore_errors)
        elif endpoint == "stem_acquisition_param":
            ignore_errors = bool(query.get("ignore_errors", [False])[0])
            response = self.get_microscope().set_stem_acquisition_param(decoded_content, ignore_errors=ignore_errors)
        elif endpoint.startswith("detector_param/"):
            name = unquote(endpoint[15:])
            response = self.get_microscope().set_detector_param(name, decoded_content)
        elif endpoint == "normalize":
            self.get_microscope().normalize(decoded_content)
        elif endpoint == "column_valves_open":
            state = bool(decoded_content)
            assert isinstance(self.server, MicroscopeServer)
            self.get_microscope().set_column_valves_open(state)
        else:
            raise KeyError("Unknown endpoint: '%s'" % endpoint)
        return response

    # Handler for the GET requests
    def do_GET(self):
        try:
            request = urlparse(self.path)
            if request.path.startswith("/v1/"):
                response = self.do_GET_V1(request.path[4:], parse_qs(request.query))
            else:
                raise KeyError('Unknown API version: %s' % self.path)
        except KeyError as exc:
            self.log_error("KeyError raised during handling of GET request '%s': %s" % (self.path, repr(exc)))
            self.send_error(404, str(exc))
        except Exception as exc:
            self.log_error("Exception raised during handling of GET request '%s': %s" % (self.path, repr(exc)))
            self.send_error(500, "Error handling request '%s': %s" % (self.path, str(exc)))
        else:
            self.build_response(response)

    # Handler for the PUT requests
    def do_PUT(self):
        try:
            request = urlparse(self.path)
            if request.path.startswith("/v1/"):
                response = self.do_PUT_V1(request.path[4:], parse_qs(request.query))
            else:
                raise KeyError('Unknown API version: %s' % self.path)
        except KeyError as exc:
            self.log_error("KeyError raised during handling of GET request '%s': %s" % (self.path, repr(exc)))
            self.send_error(404, str(exc))
        except Exception as exc:
            self.log_error("Exception raised during handling of GET request '%s': %s" % (self.path, repr(exc)))
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

    #if platform.system() != "Windows":
    #    raise NotImplementedError("This server should be started on the microscope PC (Windows only)")

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
