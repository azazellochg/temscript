#!/usr/bin/env python3
from setuptools import setup
# To use a consistent encoding
from codecs import open
from os import path

import pytemscript
from pytemscript import __version__

here = path.abspath(path.dirname(__file__))

# Long description
with open(path.join(here, "README.md"), "r", encoding="utf-8") as fp:
    long_description = fp.read()

setup(name='pytemscript',
      version=__version__,
      description='TEM Scripting adapter for FEI/TFS microscopes',
      author='Tore Niermann, Grigory Sharov',
      author_email='tore.niermann@tu-berlin.de, gsharov@mrc-lmb.cam.ac.uk',
      long_description=long_description,
      long_description_content_type='text/markdown',
      packages=['pytemscript'],
      platforms=['any'],
      license="GNU General Public License v3 (GPLv3)",
      python_requires='>=3.4',
      classifiers=[
          'Development Status :: 3 - Alpha',
          'Intended Audience :: Science/Research',
          'Intended Audience :: Developers',
          'Operating System :: OS Independent',
          'Programming Language :: Python :: 3',
          'Topic :: Scientific/Engineering',
          'Topic :: Software Development :: Libraries',
          'License :: OSI Approved :: GNU General Public License v3 (GPLv3)',
          'Operating System :: OS Independent'
      ],
      install_requires=['numpy', 'comtypes'],
      entry_points={'console_scripts': ['pytemscript-server = pytemscript.server:run_server']},
      url="https://github.com/azazellochg/pytemscript",
      project_urls={
          "Source": "https://github.com/azazellochg/pytemscript",
          # 'Documentation': "https://pytemscript.readthedocs.io/"
      }
      )
