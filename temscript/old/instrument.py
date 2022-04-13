import platform
if platform.system() == "Windows":
    from temscript.com.instrument_com import *
    from temscript.com.instrument_adv_com import *
else:
    from temscript.com.instrument_stubs import *


__all__ = ('GetInstrument', 'GetAdvancedInstrument')
