.. _microscope:

Microscope class
================

The :class:`Microscope` class provides a Python interface to the microscope.
Below are the main class properties, each represented by a separate class:

    * acquisition = :meth:`~pytemscript.modules.Acquisition`
    * apertures = :meth:`~pytemscript.modules.Apertures`
    * autoloader = :meth:`~pytemscript.modules.Autoloader`
    * detectors = :meth:`~pytemscript.modules.Detectors`
    * energy_filter = :meth:`~pytemscript.modules.EnergyFilter`
    * gun = :meth:`~pytemscript.modules.Gun`
    * optics = :meth:`~pytemscript.modules.Optics`

        * illumination = :meth:`~pytemscript.modules.Illumination`
        * projection = :meth:`~pytemscript.modules.Projection`

    * piezo_stage = :meth:`~pytemscript.modules.PiezoStage`
    * stage = :meth:`~pytemscript.modules.Stage`
    * stem = :meth:`~pytemscript.modules.Stem`
    * temperature = :meth:`~pytemscript.modules.Temperature`
    * user_buttons = :meth:`~pytemscript.modules.UserButtons`
    * user_door = :meth:`~pytemscript.modules.UserDoor`
    * vacuum = :meth:`~pytemscript.modules.Vacuum`

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


Image object
------------

Two acquisition functions: :meth:`~pytemscript.modules.Acquisition.acquire_tem_image` and
:meth:`~pytemscript.modules.Acquisition.acquire_stem_image` return an :class:`Image` object
that has the following methods and properties:

.. autoclass:: pytemscript.modules.Image
    :members: width, height, bit_depth, pixel_type, data, save, name, metadata

Documentation
-------------

.. automodule:: pytemscript.modules
    :members: Acquisition, Apertures, Autoloader, Detectors, EnergyFilter, Gun, Optics, Illumination, Projection, PiezoStage, Stage, Stem, Temperature, UserButtons, UserDoor, Vacuum
