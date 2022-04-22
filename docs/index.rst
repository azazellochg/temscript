Introduction
------------

The ``pytemscript`` package provides a Python wrapper for both standard and advanced scripting
interfaces of Thermo Fisher Scientific and FEI microscopes. The functionality is
limited to the functionality of the original scripting interfaces. For detailed information
about TEM scripting see the documentation accompanying your microscope.

The ``pytemscript`` package provides two interfaces to the microscope. The first one
corresponds directly to the COM interface. The other interface is a more high level interface.
Within the ``pytemscript`` package three implementation for the high level interface are provided,
one for running scripts directly on the microscope PC, one to run scripts remotely over network, and
finally a dummy implementation for offline development & testing exists.

Currently the ``pytemscript`` package requires Python 3.4 or higher. The current plan is to keep the minimum
supported Python version at 3.4, since this is the latest Python version supporting Windows XP.

This is a GPL fork of the original BSD-licensed project: https://github.com/niermann/temscript
New changes and this whole product is distributed under either version 3 of the GPL License, or
(at your option) any later version.

Documentation
-------------

The documentation can be found at https://pytemscript.readthedocs.io

Installation
------------

.. warning:: The project is still in development phase, no beta version has been released yet. Installing from sources is recommended.

Requirements:

    * python >= 3.4
    * comtypes
    * mrcfile and numpy (optional, only to save images in mrc format)
    * sphinx-rtd-theme (optional, only for building documentation)
    * matplotlib (optional, only for running tests)

Installation from PyPI on Windows
#################################

This assumes you have connection to the internet.

Execute from the command line (assuming you have your Python interpreter in the path):

.. code-block:: python

    py -m pip install --upgrade pip
    py -m pip install pytemscript

Offline-Installation from wheels file on Windows
################################################

This assumes you have downloaded the wheels file <downloaded-wheels-file>.whl for temscript and comtypes into the current folder.

Execute from the command line (assuming you have your Python interpreter in the path:

.. code-block:: python

    py -m pip install numpy comtypes pytemscript --no-index --find-links .

If you want to install pytemscript from sources (you still need to download comtypes *.whl):

.. code-block:: python

    py -m pip install numpy comtypes --no-index --find-links .
    py -m pip install -e <source_directory>

Supported functions of the COM interface
----------------------------------------

Relative to Titan V1.9 standard scripting adapter:

    * Acquisition: complete
    * ApertureMechanismCollection: complete, untested (requires a separate license)
    * AutoLoader: complete
    * BlankerShutter: complete
    * Camera: complete
    * Configuration: complete
    * Gun: complete
    * Gun1: incomplete, untested (requires TEM Server 7.10)
    * Illumination: complete
    * InstrumentModeControl: complete
    * Projection: complete
    * Stage: complete
    * TemperatureControl: complete
    * UserButton(s): complete, no events handling
    * Vacuum: complete

Relative to Titan V1.2 advanced scripting adapter:

    * Acquisitions: complete
    * Phaseplate: complete
    * PiezoStage: complete (untested)
    * Source: complete (untested)
    * UserDoorHatch: complete (untested)

Quick example
-------------

Execute this on the microscope PC (with ``pytemscript`` package installed) to create an instance of the local
:class:`Microscope` interface:

.. code-block:: python

    from pytemscript.microscope import Microscope
    microscope = Microscope()

Show the current acceleration voltage:

.. code-block:: python

    microscope.gun.voltage
    300.0

Move beam:

.. code-block:: python

    beam_pos = microscope.optics.illumination.beam_shift
    print(beam_pos)
    (0.0, 0.0)
    new_beam_pos = beam_pos[0], beam_pos[1] + 1.02
    microscope.optics.illumination.beam_shift(new_beam_pos)

Take an image:

.. code-block:: python

    image = microscope.acquisition.acquire_tem_image("BM-Ceta",
                                                     size=AcqImageSize.FULL,  # <-- see enumerations
                                                     exp_time=0.5,
                                                     binning=2)
    image.save("img.mrc")

Disclaimer
----------

Copyright (c) 2012-2021 by Tore Niermann
Contact: tore.niermann (at) tu-berlin.de

Copyleft 2022 by Grigory Sharov
Contact: gsharov (at) mrc-lmb.cam.ac.uk

All product and company names are trademarks or registered trademarks
of their respective holders. Use of them does not imply any affiliation
with or endorsement by them.

This program is free software; you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation; either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`

