#!/usr/bin/env python3

from temscript.microscope import Microscope
from temscript.utils.enums import *


def test_projection(microscope, eftem=False):
    print("Testing projection...")
    projection = microscope.optics.projection
    print("Mode:", projection.mode)
    print("Focus:", projection.focus)
    print("Magnification:", projection.magnification)
    print("MagnificationIndex:", projection.magnificationIndex)
    print("CameraLengthIndex:", projection.camera_length_index)
    print("ImageShift:", projection.image_shift)
    print("ImageBeamShift:", projection.image_beam_shift)
    print("DiffractionShift:", projection.diffraction_shift)
    print("DiffractionStigmator:", projection.diffraction_stigmator)
    print("ObjectiveStigmator:", projection.objective_stigmator)
    print("SubMode:", projection.magnification_range)
    print("LensProgram:", projection.is_eftem_on)
    print("ImageRotation:", projection.image_rotation)
    # print("DetectorShift:", projection.DetectorShift)
    # print("DetectorShiftMode:", projection.DetectorShiftMode)
    print("ImageBeamTilt:", projection.image_beam_tilt)
    print("LensProgram:", projection.is_eftem_on)

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
        print("Image name:", image.name)
        print("Image size:", image.width, image.height)
        print("Bit depth:", image.bit_depth)

        if image.metadata is not None:
            print("Binning:", image.metadata['Binning.Width'])
            print("Exp time:", image.metadata['ExposureTime'])
            print("Timestamp:", image.metadata['TimeStamp'])

        fn = cam_name + ".mrc"
        print(f"Saving to {fn}")
        image.save(filename=fn, normalize=False)

    for det in detectors:
        image = acquisition.acquire_stem_image(det,
                                               size=AcqImageSize.FULL,
                                               dwell_time=1e-5,
                                               binning=2)
        fn = det + ".mrc"
        print(f"Saving to {fn}")
        image.save(filename=fn, normalize=False)


def test_vacuum(microscope, full_test=False):
    print("Testing vacuum...")
    vacuum = microscope.vacuum
    print("Status:", vacuum.status)
    print("PVPRunning:", vacuum.is_buffer_running)
    print("ColumnValvesOpen:", vacuum.is_colvalves_open)
    print("Gauges:", vacuum.gauges)

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
    print("Status:", stage.status)
    print("Position:", pos)
    print("Holder:", stage.holder_type)
    print("Limits:", stage.limits)

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


def test_cameras(microscope):
    print("Testing cameras...")
    dets = microscope.detectors
    print("Film settings:", dets.film_settings)
    print("Cameras:", dets.cameras)
    print("STEM detectors:", dets.stem_detectors)


def test_optics(microscope):
    print("Testing optics...")
    opt = microscope.optics
    print("ScreenCurrent:", opt.screen_current)
    print("BeamBlanked:", opt.is_beam_blanked)
    print("AutoNormalizeEnabled:", opt.is_autonormalize_on)
    print("ShutterOverrideOn:", opt.is_shutter_override_on)
    opt.beam_blank()
    opt.beam_unblank()
    opt.normalize(ProjectionNormalization.OBJECTIVE)
    opt.normalize_all()


def test_illumination(microscope):
    print("Testing illumination...")
    illum = microscope.optics.illumination
    print("Mode:", illum.mode)
    print("SpotsizeIndex:", illum.spotsize)
    print("Intensity:", illum.intensity)
    print("IntensityZoomEnabled:", illum.intensity_zoom)
    print("IntensityLimitEnabled:", illum.intensity_limit)
    print("Shift:", illum.beam_shift)
    print("Tilt:", illum.beam_tilt)
    print("RotationCenter:", illum.rotation_center)
    print("CondenserStigmator:", illum.condenser_stigmator)
    print("DFMode:", illum.dark_field_mode)

    if microscope.condenser_system == CondenserLensSystem.THREE_CONDENSER_LENSES:
        print("CondenserMode:", illum.condenser_mode)
        print("IlluminatedArea:", illum.illuminated_area)
        print("ProbeDefocus:", illum.probe_defocus)
        print("ConvergenceAngle:", illum.convergence_angle)
        print("C3ImageDistanceParallelOffset:", illum.C3ImageDistanceParallelOffset)


def test_stem(microscope):
    print("Testing STEM...")
    stem = microscope.stem
    print("StemAvailable:", stem.is_stem_available)

    if stem.is_stem_available:
        stem.enable()
        print("Illumination.StemMagnification:", stem.stem_magnification)
        print("Illumination.StemRotation:", stem.stem_rotation)
        print("Illumination.StemFullScanFieldOfView:", stem.stem_scan_fov)
        stem.disable()


def test_gun(microscope):
    print("Testing gun...")
    gun = microscope.gun
    print("HTState:", gun.ht_state)
    print("HTValue:", gun.voltage)
    print("HTMaxValue:", gun.voltage_max)
    print("Shift:", gun.shift)
    print("Tilt:", gun.tilt)


def test_apertures(microscope):
    print("Testing apertures...")
    aps = microscope.apertures
    print("GetCurrentPresetPosition", aps.vpp_position)
    aps.vpp_next_position()


def test_general(microscope, check_door=False):
    print("Testing configuration...")

    print("Configuration.ProductFamily:", microscope.family)
    print("UserButtons:", microscope.user_buttons)
    print("BlankerShutter.ShutterOverrideOn:",
          microscope.optics.is_shutter_override_on)
    print("Condenser system:", microscope.condenser_system)
    print("Licenses:", microscope.check_license())

    if check_door:
        print("User door:", microscope.user_door.state)


if __name__ == '__main__':
    print("Starting Test...")

    full_test = False
    microscope = Microscope()
    test_projection(microscope, eftem=True)
    test_cameras(microscope)
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
