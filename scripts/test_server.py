#!/usr/bin/env python3
import threading

import pytemscript.utils.server as temserver
from pytemscript import RemoteMicroscope


def create_server(*args):
    temserver.main(args)


if __name__ == '__main__':
    print("Starting remote server test...")
    passed = "FAILED"

    proc = threading.Thread(target=create_server,
                            args=["-p 8080"])
    proc.start()

    client = RemoteMicroscope(port=8080)
    try:
        assert client._request("GET", "/has/_tem") is True
        assert client._request("GET", "/get/_tem.Illumination.Shift.X")[1] is not None
        passed = "PASSED"
    except:
        pass

    print("Test %s! Press Ctrl+C" % passed)
