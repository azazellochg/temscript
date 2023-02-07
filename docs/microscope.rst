.. _microscope:

Microscope class
================

The :class:`Microscope` class provides a Python interface to the microscope.
Below are the main class properties, each represented by a separate class:

    * acquisition = :meth:`~pytemscript.microscope.Acquisition`
    * apertures = :meth:`~pytemscript.microscope.Apertures`
    * autoloader = :meth:`~pytemscript.microscope.Autoloader`
    * detectors = :meth:`~pytemscript.microscope.Detectors`
    * energy_filter = :meth:`~pytemscript.microscope.EnergyFilter`
    * gun = :meth:`~pytemscript.microscope.Gun`
    * optics = :meth:`~pytemscript.microscope.Optics`

        * illumination = :meth:`~pytemscript.microscope.Illumination`
        * projection = :meth:`~pytemscript.microscope.Projection`

    * piezo_stage = :meth:`~pytemscript.microscope.PiezoStage`
    * stage = :meth:`~pytemscript.microscope.Stage`
    * stem = :meth:`~pytemscript.microscope.Stem`
    * temperature = :meth:`~pytemscript.microscope.Temperature`
    * user_door = :meth:`~pytemscript.microscope.UserDoor`
    * vacuum = :meth:`~pytemscript.microscope.Vacuum`


Image object
------------

Two acquisition functions: :meth:`~pytemscript.microscope.Acquisition.acquire_tem_image` and
:meth:`~pytemscript.microscope.Acquisition.acquire_stem_image` return an :class:`Image` object
that has the following methods and properties:

.. autoclass:: pytemscript.microscope.Image
    :members: width, height, bit_depth, pixel_type, data, save, name, metadata


Example usage
-------------

.. code-block:: python

    microscope = Microscope()
    curr_pos = microscope.stage.position
    print(curr_pos['Y'])
    24.05
    microscope.stage.move_to(x=-30, y=25.5)

    beam_shift = microscope.optics.illumination.beam_shift
    defocus = microscope.optics.projection.defocus
    microscope.optics.normalize_all()

Documentation
-------------

.. automodule:: pytemscript.microscope
    :members:
    :exclude-members: Image
