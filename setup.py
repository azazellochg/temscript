#!/usr/bin/env python3
from setuptools import setup
# To use a consistent encoding
from codecs import open
from os import path

import temscript
from temscript import __version__

here = path.abspath(path.dirname(__file__))

# Long description
with open(path.join(here, "README.md"), "r", encoding="utf-8") as fp:
    long_description = fp.read()

setup(name='temscript',
      version=__version__,
      description='TEM Scripting adapter for FEI microscopes',
      author='Tore Niermann, Grigory Sharov',
      author_email='tore.niermann@tu-berlin.de, gsharov@mrc-lmb.cam.ac.uk',
      long_description=long_description,
      long_description_content_type='text/markdown',
      packages=['temscript'],
      platforms=['any'],
      license="BSD 3-Clause License",
      python_requires='>=3.4',
      classifiers=[
          'Development Status :: 4 - Beta',
          'Intended Audience :: Science/Research',
          'Intended Audience :: Developers',
          'Operating System :: OS Independent',
          'Programming Language :: Python :: 3',
          'Topic :: Scientific/Engineering',
          'Topic :: Software Development :: Libraries',
          'License :: OSI Approved :: BSD License',
          'Operating System :: OS Independent'
      ],
      install_requires=['numpy'],
      entry_points={'console_scripts': ['temscript-server = temscript.server:run_server']},
      url="https://github.com/azazellochg/temscript",
      project_urls={
          "Source": "https://github.com/azazellochg/temscript",
          'Documentation': "https://temscript.readthedocs.io/"
      }
      )
