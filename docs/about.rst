About
=====

The COM interface
-----------------

The methods and classes represent the COM objects exposed by the *Scripting* interface.
The interface is described in detail in the scripting manual of your microscope
(usually in the file ``scripting.pdf`` located in the ``C:\Titan\Tem_help\manual`` or
``C:\Tecnai\tem_help\manual`` directories). Advanced scripting manual can be found in
``C:\Titan\Scripting\Advanced TEM Scripting User Guide.pdf``.

The manual is your ultimate reference, this documentation will only describe the
python wrapper to the COM interface.

Microscope class
----------------

The :ref:`microscope` class provides the main interface to the microscope.

Enumerations
------------

Many of the attributes return values from enumerations. The complete list can be found in the :ref:`enumerations` section.

.. versionchanged:: 2.0
    All methods of the COM interface now directly return the enumeration objects. The constants
    from pytemscript version 1.x are not defined anymore. The numerical values still can be accessed
    by querying the corresponding enum, e.g. ``psmSA`` corresponds to ``ProjectionSubMode.SA``.

Vectors
-------

Some object attributes handle two dimensional vectors (e.g. ``ImageShift``). These
attributes return ``(x, y)`` tuples and expect iterable objects (``tuple``,
``list``, ...) with two floats when written (numpy arrays with two entries also work).

.. code-block:: python

    beam_pos = microscope.optics.illumination.beam_shift
    print(beam_pos)
    (0.0, 0.0)
    new_beam_pos = beam_pos[0], beam_pos[1] + 1e-6
    microscope.optics.illumination.beam_shift(new_beam_pos)
