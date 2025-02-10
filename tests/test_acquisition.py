#!/usr/bin/env python3
import numpy as np
import math
import matplotlib.pyplot as plt

from pytemscript.microscope import Microscope
from pytemscript.utils.enums import *


def print_stats(cam_name, image, binning, time):
    metadata = image.metadata
    if metadata is not None:
        print("\tTimestamp: ", metadata['TimeStamp'])
        assert int(metadata['Binning.Width']) == binning
        assert math.isclose(float(metadata['ExposureTime']), time, abs_tol=0.01)

    print("\tBit depth: ", image.bit_depth)
    print("\tSize: ", image.data.shape[1], image.data.shape[0])
    print("\tMean: ", np.mean(image.data))
    vmin = np.percentile(image.data, 3)
    vmax = np.percentile(image.data, 97)
    print("\tStdDev: ", np.std(image.data))
    plt.imshow(image.data, interpolation="nearest", cmap="gray",
               vmin=vmin, vmax=vmax)
    print("\tStdDev: ", np.std(image.data))
    plt.colorbar()
    plt.suptitle(cam_name)
    plt.ion()
    plt.show()
    plt.pause(1.0)


def camera_acquire(microscope, cam_name, exp_time, binning, **kwargs):
    """ Acquire a test image and check output metadata. """

    image = microscope.acquisition.acquire_tem_image(cam_name,
                                                     size=AcqImageSize.FULL,
                                                     exp_time=exp_time,
                                                     binning=binning,
                                                     **kwargs)
    print_stats("TEM camera: " + cam_name, image, binning, exp_time)


def detector_acquire(microscope, cam_name, dwell_time, binning, **kwargs):
    image = microscope.acquisition.acquire_stem_image(cam_name,
                                                      size=AcqImageSize.FULL,
                                                      dwell_time=dwell_time,
                                                      binning=binning,
                                                      **kwargs)
    print_stats("STEM detector: " + cam_name, image, binning, dwell_time)


if __name__ == '__main__':
    print("Starting acquisition test...")

    microscope = Microscope()
    cameras = microscope.detectors.cameras
    print("Available detectors:\n", cameras)

    if "BM-Ceta" in cameras:
        camera_acquire(microscope, "BM-Ceta", exp_time=1, binning=2)
    if "BM-Falcon" in cameras:
        camera_acquire(microscope, "BM-Falcon", exp_time=0.5, binning=2)
        camera_acquire(microscope, "BM-Falcon", exp_time=3, binning=1,
                       align_image=True, electron_counting=True,
                       frame_ranges=[(1, 2), (2, 3)])

    if microscope.stem.is_available:
        microscope.stem.enable()
        detectors = microscope.detectors.stem_detectors
        if "BF" in detectors:
            detector_acquire(microscope, "BF", dwell_time=1e-5, binning=2)
        microscope.stem.disable()
