.. _microscope:

Microscope class
================

The :class:`Microscope` class provides a Python interface to the microscope.
Below are the main class properties, each represented by a separate class:

    * acquisition = :meth:`~temscript.microscope.Acquisition`
    * apertures = :meth:`~temscript.microscope.Apertures`
    * autoloader = :meth:`~temscript.microscope.Autoloader`
    * detectors = :meth:`~temscript.microscope.Detectors`
    * gun = :meth:`~temscript.microscope.Gun`
    * optics = :meth:`~temscript.microscope.Optics`

        * illumination = :meth:`~temscript.microscope.Illumination`
        * projection = :meth:`~temscript.microscope.Projection`

    * piezo_stage = :meth:`~temscript.microscope.PiezoStage`
    * stage = :meth:`~temscript.microscope.Stage`
    * stem = :meth:`~temscript.microscope.Stem`
    * temperature = :meth:`~temscript.microscope.Temperature`
    * user_door = :meth:`~temscript.microscope.UserDoor`
    * vacuum = :meth:`~temscript.microscope.Vacuum`


Image object
------------

Two acquisition functions: :meth:`~temscript.microscope.Acquisition.acquire_tem_image` and
:meth:`~temscript.microscope.Acquisition.acquire_stem_image` return an :class:`Image` object
that has the following methods:

.. autoclass:: temscript.base_microscope.Image
    :members:


Example usage
-------------

.. code-block:: python

    microscope = Microscope()
    curr_pos = microscope.stage.position
    print(curr_pos['Y'])
    1.1e-6
    microscope.stage.move_to(x=-1e-6)

    beam_shift = microscope.optics.illumination.beam_shift


Documentation
-------------

.. automodule:: temscript.microscope
    :members:

