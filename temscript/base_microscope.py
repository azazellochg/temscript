import logging
from comtypes.client import CreateObject

from utils.constants import *


class BaseMicroscope:

    def __init__(self, address=None, timeout=None, simulate=False, logLevel=logging.INFO):
        self.tem = None
        self.tem_adv = None

        logging.basicConfig(level=logLevel,
                            handlers=[
                                logging.FileHandler("debug.log"),
                                logging.StreamHandler()])

        if simulate:
            raise NotImplementedError()
        elif address is None:
            # local connection
            self._createInstrument()
        else:
            raise NotImplementedError()

    def _createInstrument(self):
        """ Try to use both std and advanced scripting. """
        try:
            self.tem_adv = CreateObject(SCRIPTING_ADV)
            self.tem = CreateObject(SCRIPTING_STD)
            logging.info(f"Connected to {SCRIPTING_ADV} and {SCRIPTING_STD}")
        except:
            logging.info(f"Could not connect to {SCRIPTING_ADV}")
            self.tem = CreateObject(SCRIPTING_STD)
        else:
            self.tem = CreateObject(SCRIPTING_TECNAI)
            logging.info(f"Connected to {SCRIPTING_TECNAI}")
        finally:
            raise Exception("Could not connect to the instrument")
