.. _remote:

Remote server
=============

.. important:: The server is currently in development and not ready to use!

If remote scripting of the microscope is required, the pytemscript server must run on the microscope PC.
The server supports different connection methods:

 * socket-based (port 39000)
 * ZMQ (port 5555)
 * GRPC (port 50051)

If you would like to use ZMQ server, you need to install ZMQ server on the microscope PC and ``pyzmq`` on the client PC (via pip: `py -m pip install pyzmq`).

If you would like to use GRPC server, you need to install it via pip (`py -m pip install grpcio`) on both client and server PC.

.. warning::

    The server provides no means of security or authorization control itself.
    Thus it is highly recommended to let the server only listen to internal networks or at least route it through a reverse proxy, which implements sufficient security.

Running the server
------------------

The pytemscript server is started on the microscope PC by the ``pytemscript-server`` command:

.. code-block:: none

    usage: pytemscript-server [-h] [-t {socket,zmq,grpc}] [-p PORT] [--host HOST] [--useLD] [--useTecnaiCCD] [-d]

    optional arguments:
    -h, --help                      show this help message and exit
    -t, --type {socket,zmq,grpc}    Server type to use: socket, zmq or grpc (default: socket)
    -p, --port PORT                 Specify port on which the server is listening (default: 39000)
    --host HOST                     Specify host address on which the server is listening (default: 127.0.0.1)
    --useLD                         Connect to LowDose server on microscope PC (limited control only) (default: True)
    --useTecnaiCCD                  Connect to TecnaiCCD plugin on microscope PC that controls Digital Micrograph (may be faster than via TIA / std scripting) (default: False)
    -d, --debug                     Enable debug mode (default: False)

Default connection type is the socket listening on port 39000.

Connecting to the server
------------------------

The interface is essentially the same as for the local client:

.. code-block:: python

    from pytemscript.microscope import Microscope
    microscope = Microscope(connection="socket", host="127.0.0.1", port=39000)
