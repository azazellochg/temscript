# About

The ``temscript`` package provides a Python wrapper for both standard and advanced scripting
interfaces of Thermo Fisher Scientific and FEI microscopes. The functionality is
limited to the functionality of the original scripting interfaces. For detailed information
about TEM scripting see the documentation accompanying your microscope.

The ``temscript`` package provides two interfaces to the microscope. The first one
corresponds directly to the COM interface. The other interface is a more high level interface.
Within the ``temscript`` package three implementation for the high level interface are provided,
one for running scripts directly on the microscope PC, one to run scripts remotely over network, and
finally a dummy implementation for offline development & testing exists.

Currently the ``temscript`` package requires Python 3.4 or higher. The current plan is to keep the minimum
supported Python version at 3.4, since this is the latest Python version supporting Windows XP.

This is a GPL fork of the original BSD-licensed project: https://github.com/niermann/temscript
New changes and this whole product is distributed under either version 3 of the GPL License, or
(at your option) any later version.

# Documentation

The documentation can be found in the docs/ folder.

# Installation

Requirements:
* Python >= 3.4
* Numpy
* Sphinx (only for building documentation)

On all platforms the dummy and remote high level interfaces are provided. 
On Windows platforms the package provides the Python wrapper
to the scripting COM interface via comtypes library. However, trying to instantiate this wrapper
will fail if the scripting COM classes are not installed locally.

### Installation from PyPI (using pip)

This assumes you have connection to the internet. 

Execute from the command line (assuming you have your python interpreter in the path, this might require superuser or 
administrator privileges):
    
    python3 -m pip install --upgrade pip
    python3 -m pip install temscript

### Offline-Installation from wheels file (using pip)

This assumes you have downloaded the wheels file <downloaded-wheels-file>.whl 

Execute from the command line (assuming you have your python interpreter in the path, this might require superuser or 
administrator privileges):
    
    python3 -m pip install --upgrade pip
    python3 -m pip install <downloaded-wheels-file>.whl

### Installation from sources (using pip)

This assumes you have downloaded and extracted the sources into the directory <source_directory> (alternative have
cloned the sources from GitHub into <source_directory>). 

Execute from the command line (assuming you have your python interpreter in the path, this might require superuser or 
administrator privileges):
    
    python3 -m pip install --upgrade pip
    python3 -m pip install <source_directory>


# Supported functions of the COM interface

Relative to Titan V1.9 standard scripting adapter:

* Acquisition: complete
* ApertureMechanismCollection: complete
* AutoLoader: complete
* BlankerShutter: complete
* Camera: complete
* Configuration: complete
* Gun: complete
* Gun1: missing (requires TEM Server 7.10)
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
* PiezoStage: complete
* Source: complete
* UserDoorHatch: complete

# Disclaimer

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
