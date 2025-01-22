#!/usr/bin/env python3
from setuptools import setup
# To use a consistent encoding
from codecs import open
from os import path
from pytemscript import __version__

here = path.abspath(path.dirname(__file__))
# Long description
with open(path.join(here, "README.rst"), "r", encoding="utf-8") as fp:
    long_description = fp.read()

# Load requirements.txt
with open('requirements.txt') as f:
    requirements = f.read().splitlines()

setup(name='pytemscript',
      version=__version__,
      description='TEM Scripting adapter for FEI/TFS microscopes',
      author='Tore Niermann, Grigory Sharov',
      author_email='tore.niermann@tu-berlin.de, gsharov@mrc-lmb.cam.ac.uk',
      long_description=long_description,
      long_description_content_type='text/x-rst',
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
      keywords='TEM python',
      install_requires=[requirements],
      extras_require={
          "extra": ["matplotlib"]
      },
      #entry_points={'console_scripts': ['pytemscript-server = pytemscript.utils.http_server:main']},
      url="https://github.com/azazellochg/pytemscript",
      project_urls={
          "Source": "https://github.com/azazellochg/pytemscript",
          "Documentation": "https://pytemscript.readthedocs.io/"
      }
)
