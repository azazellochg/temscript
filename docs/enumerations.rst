.. _enumerations:

Enumerations
============

Enumerations are represented with :class:`IntEnum` objects and are used to describe TEM Scripting constants.
When a property returns an enumeration, it will print its **name**. When you assign a variable to an enumeration,
it will use its integer **value**. To see the accepted values, you can switch to source code using links below.

Example:

.. code-block:: python

    from pytemscript.microscope import Microscope
    from pytemscript.utils.enums import *

    microscope = Microscope()
    stage = microscope.stage
    print(stage.status)
    'READY'  # <-- enumeration name

    camera_size = AcqImageSize.FULL  # <-- enumeration value
    image = microscope.acquisition.acquire_tem_image("BM-Ceta",
                                                     size=camera_size,
                                                     exp_time=0.5,
                                                     binning=2)

.. automodule:: pytemscript.utils.enums
    :members:
