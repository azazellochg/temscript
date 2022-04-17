.. _microscope:

Microscope class
================

The :class:`Microscope` class provides a Python interface to the microscope.
Below are the main class properties, each represented by a separate class:

    * acquisition = :meth:`~temscript.microscope.Acquisition`
    * detectors = :meth:`~temscript.microscope.Detectors`
    * gun = :meth:`~temscript.microscope.Gun`
    * optics = :meth:`~temscript.microscope.Optics`

        * illumination = :meth:`~temscript.microscope.Illumination`
        * projection = :meth:`~temscript.microscope.Projection`

    * stem = :meth:`~temscript.microscope.Stem`
    * apertures = :meth:`~temscript.microscope.Apertures`
    * temperature = :meth:`~temscript.microscope.Temperature`
    * vacuum = :meth:`~temscript.microscope.Vacuum`
    * autoloader = :meth:`~temscript.microscope.Autoloader`
    * stage = :meth:`~temscript.microscope.Stage`
    * piezo_stage = :meth:`~temscript.microscope.PiezoStage`
    * user_door = :meth:`~temscript.microscope.UserDoor`

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

