#!/usr/bin/env python3

from pytemscript.microscope import Microscope
from pytemscript.utils.enums import *


def test_projection(microscope, eftem=False):
    print("Testing projection...")
    projection = microscope.optics.projection
    print("\tMode:", projection.mode)
    print("\tFocus:", projection.focus)
    print("\tMagnification:", projection.magnification)
    print("\tMagnificationIndex:", projection.magnificationIndex)
    print("\tCameraLengthIndex:", projection.camera_length_index)
    print("\tImageShift:", projection.image_shift)
    print("\tImageBeamShift:", projection.image_beam_shift)
    print("\tDiffractionShift:", projection.diffraction_shift)
    print("\tDiffractionStigmator:", projection.diffraction_stigmator)
    print("\tObjectiveStigmator:", projection.objective_stigmator)
    print("\tSubMode:", projection.magnification_range)
    print("\tLensProgram:", projection.is_eftem_on)
    print("\tImageRotation:", projection.image_rotation)
    # print("\tDetectorShift:", projection.DetectorShift)
    # print("\tDetectorShiftMode:", projection.DetectorShiftMode)
    print("\tImageBeamTilt:", projection.image_beam_tilt)
    print("\tLensProgram:", projection.is_eftem_on)

    projection.reset_defocus()

    if eftem:
        projection.eftem_on()
        projection.eftem_off()


def test_acquisition(microscope):
    print("Testing acquisition...")
    acquisition = microscope.acquisition
    cameras = microscope.detectors.cameras
    detectors = microscope.detectors.stem_detectors

    for cam_name in cameras:
        image = acquisition.acquire_tem_image(cam_name,
                                              size=AcqImageSize.FULL,
                                              exp_time=0.25,
                                              binning=2)
        print("\tImage name:", image.name)
        print("\tImage size:", image.width, image.height)
        print("\tBit depth:", image.bit_depth)

        if image.metadata is not None:
            print("\tBinning:", image.metadata['Binning.Width'])
            print("\tExp time:", image.metadata['ExposureTime'])
            print("\tTimestamp:", image.metadata['TimeStamp'])

        fn = cam_name + ".mrc"
        print("Saving to ", fn)
        image.save(filename=fn, normalize=False)

    for det in detectors:
        image = acquisition.acquire_stem_image(det,
                                               size=AcqImageSize.FULL,
                                               dwell_time=1e-5,
                                               binning=2)
        fn = det + ".mrc"
        print("Saving to ", fn)
        image.save(filename=fn, normalize=False)


def test_vacuum(microscope, full_test=False):
    print("Testing vacuum...")
    vacuum = microscope.vacuum
    print("\tStatus:", vacuum.status)
    print("\tPVPRunning:", vacuum.is_buffer_running)
    print("\tColumnValvesOpen:", vacuum.is_colvalves_open)
    print("\tGauges:", vacuum.gauges)

    if full_test:
        vacuum.colvalves_open()
        vacuum.colvalves_close()
        vacuum.run_buffer_cycle()


def test_temperature(microscope):
    print("Testing TemperatureControl...")
    temp = microscope.temperature
    print("\tRefrigerantLevel (autoloader):",
          temp.dewar_level(RefrigerantDewar.AUTOLOADER_DEWAR))
    print("\tRefrigerantLevel (column):",
          temp.dewar_level(RefrigerantDewar.COLUMN_DEWAR))
    print("\tDewarsRemainingTime:", temp.dewars_remaining_time)
    print("\tDewarsAreBusyFilling:", temp.is_dewars_filling)


def test_autoloader(microscope, full_test=False, slot=1):
    print("Testing Autoloader...")
    al = microscope.autoloader
    print("\tNumberOfCassetteSlots", al.number_of_cassette_slots)
    print("\tSlotStatus", al.get_slot_status(3))

    if full_test:
        al.run_inventory()
        al.load_cartridge(slot)
        al.unload_cartridge(slot)


def test_stage(microscope, do_move=False):
    print("Testing stage...")
    stage = microscope.stage
    pos = stage.position
    print("\tStatus:", stage.status)
    print("\tPosition:", pos)
    print("\tHolder:", stage.holder_type)
    print("\tLimits:", stage.limits)

    if not do_move:
        return

    print("Testing stage movement...")
    print("\tGoto(x=1e-6, y=-1e-6)")
    stage.go_to(x=1e-6, y=-1e-6)
    print("\tPosition:", stage.position)
    print("\tGoto(x=-1e-6, speed=0.5)")
    stage.go_to(x=-1e-6, speed=0.5)
    print("\tPosition:", stage.position)
    print("\tMoveTo() to original position")
    stage.move_to(**pos)
    print("\tPosition:", stage.position)


def test_detectors(microscope):
    print("Testing cameras...")
    dets = microscope.detectors
    print("\tFilm settings:", dets.film_settings)
    print("\tCameras:", dets.cameras)
    print("\tSTEM detectors:", dets.stem_detectors)


def test_optics(microscope):
    print("Testing optics...")
    opt = microscope.optics
    print("\tScreenCurrent:", opt.screen_current)
    print("\tBeamBlanked:", opt.is_beam_blanked)
    print("\tAutoNormalizeEnabled:", opt.is_autonormalize_on)
    print("\tShutterOverrideOn:", opt.is_shutter_override_on)
    opt.beam_blank()
    opt.beam_unblank()
    opt.normalize(ProjectionNormalization.OBJECTIVE)
    opt.normalize_all()


def test_illumination(microscope):
    print("Testing illumination...")
    illum = microscope.optics.illumination
    print("\tMode:", illum.mode)
    print("\tSpotsizeIndex:", illum.spotsize)
    print("\tIntensity:", illum.intensity)
    print("\tIntensityZoomEnabled:", illum.intensity_zoom)
    print("\tIntensityLimitEnabled:", illum.intensity_limit)
    print("\tShift:", illum.beam_shift)
    print("\tTilt:", illum.beam_tilt)
    print("\tRotationCenter:", illum.rotation_center)
    print("\tCondenserStigmator:", illum.condenser_stigmator)
    print("\tDFMode:", illum.dark_field_mode)

    if microscope.condenser_system == CondenserLensSystem.THREE_CONDENSER_LENSES:
        print("\tCondenserMode:", illum.condenser_mode)
        print("\tIlluminatedArea:", illum.illuminated_area)
        print("\tProbeDefocus:", illum.probe_defocus)
        print("\tConvergenceAngle:", illum.convergence_angle)
        print("\tC3ImageDistanceParallelOffset:", illum.C3ImageDistanceParallelOffset)


def test_stem(microscope):
    print("Testing STEM...")
    stem = microscope.stem
    print("\tStemAvailable:", stem.is_stem_available)

    if stem.is_stem_available:
        stem.enable()
        print("\tIllumination.StemMagnification:", stem.stem_magnification)
        print("\tIllumination.StemRotation:", stem.stem_rotation)
        print("\tIllumination.StemFullScanFieldOfView:", stem.stem_scan_fov)
        stem.disable()


def test_gun(microscope):
    print("Testing gun...")
    gun = microscope.gun
    print("\tHTState:", gun.ht_state)
    print("\tHTValue:", gun.voltage)
    print("\tHTMaxValue:", gun.voltage_max)
    print("\tShift:", gun.shift)
    print("\tTilt:", gun.tilt)


def test_apertures(microscope):
    print("Testing apertures...")
    aps = microscope.apertures
    print("\tGetCurrentPresetPosition", aps.vpp_position)
    aps.vpp_next_position()


def test_general(microscope, check_door=False):
    print("Testing configuration...")

    print("\tConfiguration.ProductFamily:", microscope.family)
    print("\tUserButtons:", microscope.user_buttons)
    print("\tBlankerShutter.ShutterOverrideOn:",
          microscope.optics.is_shutter_override_on)
    print("\tCondenser system:", microscope.condenser_system)

    if check_door:
        print("\tUser door:", microscope.user_door.state)


if __name__ == '__main__':
    print("Starting tests...")

    full_test = False
    microscope = Microscope()
    test_projection(microscope, eftem=True)
    test_detectors(microscope)
    test_vacuum(microscope, full_test=full_test)
    test_autoloader(microscope, full_test=full_test, slot=1)
    test_temperature(microscope)
    test_stage(microscope, do_move=full_test)
    test_optics(microscope)
    test_illumination(microscope)
    test_gun(microscope)
    test_general(microscope, check_door=full_test)

    if full_test:
        test_acquisition(microscope)
        test_stem(microscope)
        test_apertures(microscope)
