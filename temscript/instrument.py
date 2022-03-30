import platform
if platform.system() == "Windows":
    from ._instrument_com import *
    from ._instrument_adv_com import *
else:
    from ._instrument_stubs import *


__all__ = ('GetInstrument', 'GetAdvancedInstrument')
