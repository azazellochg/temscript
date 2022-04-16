#!/usr/bin/env python3
import numpy as np
import math
import matplotlib.pyplot as plt

from temscript.microscope import Microscope
from temscript.utils.enums import *


def camera_acquire(cam_name, exp_time, binning, **kwargs):
    """ Acquire a test image and check output metadata. """
    microscope = Microscope()

    print(f"Available detectors:\n{microscope.detectors.cameras}")

    image = microscope.acquisition.acquire_tem_image(cam_name,
                                                     size=AcqImageSize.FULL,
                                                     exp_time=exp_time,
                                                     binning=binning,
                                                     **kwargs)
    print(f"Bit depth = {image.bit_depth}")

    metadata = image.metadata
    print(f"Timestamp = {metadata['TimeStamp']}")

    assert int(metadata['Binning.Width']) == binning
    assert math.isclose(float(metadata['ExposureTime']), exp_time, abs_tol=0.01)

    print(f"\tSize: {image.data.shape[1]}, {image.data.shape[0]}")
    print(f"\tMean: {np.mean(image.data)}")
    vmin = np.percentile(image.data, 3)
    vmax = np.percentile(image.data, 97)
    print(f"\tStdDev: {np.std(image.data)}")
    plt.imshow(image.data, interpolation="nearest", cmap="gray",
               vmin=vmin, vmax=vmax)
    print(f"\tStdDev: {np.std(image.data)}")
    plt.colorbar()
    plt.suptitle(cam_name)
    plt.show()


if __name__ == '__main__':
    print("Starting Test...")

    camera_acquire("BM-Falcon", exp_time=0.5, binning=2)
    camera_acquire("BM-Falcon", exp_time=3, binning=1,
                   align_image=True, electron_counting=True,
                   frame_ranges=[(1,2), (2,3)])
