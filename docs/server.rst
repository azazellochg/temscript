.. _server:

The pytemscript server
======================

If remote scripting of the microscope is required as provided via the :class:`RemoteMicroscope` class, the pytemscript
server must run on the microscope PC. The pytemscript server provides a Web-API to the interface.

.. warning::

    The server provides no means of security or authorization control itself.

    Thus it is highly recommended to let the server
    only listen to internal networks or at least route it through a reverse proxy, which implements sufficient security.

Command line
------------

The pytemscript server is started by the ``pytemscript-server`` command line script provided with the :mod:`pytemscript``
package (obviously :mod:`pytemscript` must also be installed on the microscope PC).

.. code-block:: none

    usage: pytemscript-server [-h] [-p PORT] [--host HOST] [--null]

    optional arguments:
      -h, --help            show this help message and exit
      -p PORT, --port PORT  Specify port on which the server is listening
      --host HOST           Specify host address on which the the server is listening
      --null                Use NullMicroscope instead of local microscope as backend

Python command
--------------

Alternatively the pytemscript server can also be run from within Python using :func:`run_server` function.

.. autofunction:: pytemscript.run_server

