pytemscript documentation
=======================

.. toctree::
   :maxdepth: 1

   about
   microscope
   enumerations
   changelog
   restrictions

Introduction
------------

The ``pytemscript`` package provides a Python wrapper for both standard and advanced scripting
interfaces of Thermo Fisher Scientific and FEI microscopes. The functionality is
limited to the functionality of the original scripting interfaces. For detailed information
about TEM scripting see the documentation accompanying your microscope.

.. For remote operation of the microscope the pytemscript server must run on the microscope PC. See section :ref:`server` for details.

The section :ref:`restrictions` describes some known issues with the scripting interface itself. These are restrictions
of the original scripting interface and not issues related to the ``pytemscript`` package itself.

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
    new_beam_pos = beam_pos[0], beam_pos[1] + 1e-6
    microscope.optics.illumination.beam_shift(new_beam_pos)

Take an image:

.. code-block:: python

    image = microscope.acquisition.acquire_tem_image("BM-Ceta",
                                                     size=AcqImageSize.FULL,  # <-- see enumerations
                                                     exp_time=0.5,
                                                     binning=2)

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

