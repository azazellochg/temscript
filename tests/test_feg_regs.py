#!/usr/bin/env python3

import comtypes.client

from pytemscript.utils.constants import FEG_REGISTERS


def test_fegs():
    try:
        comtypes.CoInitialize()
        obj = comtypes.client.CreateObject(FEG_REGISTERS)
    except:
        raise RuntimeError("Could not connect to %s interface" % FEG_REGISTERS)
    print("Connected to %s" % FEG_REGISTERS)
    print("Number of FEG registers: ", obj.NumberOfRegisters)

    for feg in obj.GetRegisterList():
        print("Label", feg[0])
        print("\tDateTime", feg[1].strftime("%d-%m-%Y %H:%M:%S"))
        print("\tExtrVolt", feg[2])
        print("\tPotential", feg[3])
        print("\tGunLens", feg[4])
        print("\tExcitation", feg[5])
        print("\tMode", feg[6])
        print("\tSpot", feg[7])

    print("FEG registers options: [] checked, [] enabled")
    opts = obj.GetOptions()
    print("\tFEG settings", opts[0])
    print("\tMonochromator", opts[1])
    print("\tMode", opts[2])
    print("\tCondenser", opts[3])
    print("\tMagnification", opts[4])
    print("\tDirect alignments", opts[5])
    print("\tStigmators", opts[6])
    print("\tC1 aperture", opts[7])
    print("\tC3 aperture", opts[8])
    print("\tObj aperture", opts[9])
    print("\tSA aperture", opts[10])
    print("\tCondenser stigmator", opts[11])

    feg = obj.GetRegisterList()[0]
    settings = obj.GetRegisterValues(feg[0])
    assert int(settings[15]) == int(feg[4])  # gun lens
    assert int(settings[18]) == int(feg[7])  # spot size


if __name__ == '__main__':
    print("Checking FEG registers...")
    test_fegs()
