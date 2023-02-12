r"""

The code below is a modified version of:
https://github.com/instamatic-dev/instamatic/blob/master/instamatic/camera/gatansocket3.py
https://github.com/nysbc/leginon-py3/blob/main/pyscope/gatansocket.py

BSD 3-Clause License

Copyright (c) 2021, Stef Smeets
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this
   list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
   this list of conditions and the following disclaimer in the documentation
   and/or other materials provided with the distribution.

3. Neither the name of the copyright holder nor the names of its
   contributors may be used to endorse or promote products derived from
   this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

The script adapted from [Leginon](http://emg.nysbc.org/redmine/projects/leginon/wiki/Leginon_Homepage). Leginon is licenced under the Apache License, Version 2.0. The code (`gatansocket3.py`) was converted from Python2.7 to Python3.6+ from [here](http://emg.nysbc.org/redmine/projects/leginon/repository/revisions/trunk/entry/pyscope/gatansocket.py).

It needs the SERIALEMCCD plugin to be installed in DigitalMicrograph. The relevant instructions from the [SerialEM documentation](https://bio3d.colorado.edu/SerialEM/hlp/html/setting_up_serialem.htm) are referenced below.

### Setup [1]

To connect to DigitalMicrograph through a socket interface on the same or a different computer, such as for a K2/K3 camera with SerialEM running on an FEI microscope, you need to do the following:
 - Determine the IP address of the computer running DM on the network which that computer shares with the computer running SerialEM.  If SerialEM and DM are running on the same computer, use `127.0.0.1` for the address.
 - Copy the appropriate SerialEMCDD plugin from the SerialEM_3-x-x folder to a Plugins folder on the other computer (the one running DM). Specifically:
     - If the other computer is running 64-bit Windows, copy  `SEMCCD-GMS2.31-64.dll`, `SEMCCD-GMS3.30-64.dll`, or `SEMCCD-GMS3.31-64.dll` to `C:\ProgramData\Gatan\Plugins` and rename it to `SEMCCD-GMS2-64.dll`
     - If the other computer is running GMS2 on Windows XP or Windows 7 32-bit, copy `SEMCCD-GMS2.0-32.dll` or `SEMCCD-GMS2.3-32.dll` to `C:\Program Files\Gatan\Plugins` and rename it to `SEMCCD-GMS2-32.dll`
     - If the other computer is running GMS 1, copy `SerialEMCCD.dll` to `C:\Program Files\Gatan\DigitalMicrograph\Plugins`
 - If DM and SerialEM are running on the same computer, the installer should have placed the right plugin in the right folder, but if not, follow the procedure just given.
 - On the computer running DM, define a system environment variable `SERIALEMCCD_PORT` with the value `48890` or other selected port number, as described in the section above.
 - Make sure that this port is open for communication between SerialEM and DM. If there is any possibility of either machine being exposed to the internet, do not simply turn off the firewalls; open only this specific port in the firewall, and allow access only by the other machine or the local subnet.  Even if just this one port is exposed to the world, port scanners can interfere with communication and DM function.
 - Restart DM. Note that no registration is needed for the plugin when using a socket interface.
 - If the connection does not work, debugging output can be obtained on both sides by:
     - Setting an environment variable `SERIALEMCCD_DEBUG` with the value of `1` or `2`, where `2` will give more verbose output related to the socket operations.

[1]. https://bio3d.colorado.edu/SerialEM/hlp/html/setting_up_serialem.htm

"""

import os
import socket
import logging
import numpy as np

# enum function codes as in SocketPathway.cpp
# need to match exactly both in number and order
enum_gs = [
    'GS_ExecuteScript',
    'GS_SetDebugMode',
    'GS_SetDMVersion',
    'GS_SetCurrentCamera',
    'GS_QueueScript',
    'GS_GetAcquiredImage',
    'GS_GetDarkReference',
    'GS_GetGainReference',
    'GS_SelectCamera',
    'GS_SetReadMode',
    'GS_GetNumberOfCameras',
    'GS_IsCameraInserted',
    'GS_InsertCamera',
    'GS_GetDMVersion',
    'GS_GetDMCapabilities',
    'GS_SetShutterNormallyClosed',
    'GS_SetNoDMSettling',
    'GS_GetDSProperties',
    'GS_AcquireDSImage',
    'GS_ReturnDSChannel',
    'GS_StopDSAcquisition',
    'GS_CheckReferenceTime',
    'GS_SetK2Parameters',
    'GS_ChunkHandshake',
    'GS_SetupFileSaving',
    'GS_GetFileSaveResult',
    'GS_SetupFileSaving2',
    'GS_GetDefectList',
    'GS_SetK2Parameters2',
    'GS_StopContinuousCamera',
    'GS_GetPluginVersion',
    'GS_GetLastError',
    'GS_FreeK2GainReference',
    'GS_IsGpuAvailable',
    'GS_SetupFrameAligning',
    'GS_FrameAlignResults',
    'GS_ReturnDeferredSum',
    'GS_MakeAlignComFile',
    'GS_WaitUntilReady',
    'GS_GetLastDoseRate',
    'GS_SaveFrameMdoc',
    'GS_GetDMVersionAndBuild',
    'GS_GetTiltSumProperties',
]
# lookup table of function name to function code, starting with 1
enum_gs = {x: y for (y, x) in enumerate(enum_gs, 1)}

# C "long" -> numpy "int_"
ARGS_BUFFER_SIZE = 1024
MAX_LONG_ARGS = 16
MAX_DBL_ARGS = 8
MAX_BOOL_ARGS = 8
sArgsBuffer = np.zeros(ARGS_BUFFER_SIZE, dtype=np.byte)


class Message:
    """Information packet to send and receive on the socket.

    Initialize with the sequences of args (longs, bools, doubles) and
    optional long array.
    """

    def __init__(self, longargs=[], boolargs=[], dblargs=[], longarray=[]):
        # Strings are packaged as long array using np.frombuffer(buffer,np.int_)
        # and can be converted back with longarray.tostring()
        # add final longarg with size of the longarray
        if longarray:
            longargs = list(longargs)
            longargs.append(len(longarray))

        self.dtype = [
            ('size', np.intc),
            ('longargs', np.int_, (len(longargs),)),
            ('boolargs', np.int32, (len(boolargs),)),
            ('dblargs', np.double, (len(dblargs),)),
            ('longarray', np.int_, (len(longarray),)),
        ]
        self.array = np.zeros((), dtype=self.dtype)
        self.array['size'] = self.array.data.itemsize
        self.array['longargs'] = longargs
        self.array['boolargs'] = boolargs
        self.array['dblargs'] = dblargs
        self.array['longarray'] = longarray

        # create numpy arrays for the args and array
        # self.longargs = np.asarray(longargs, dtype=np.int_)
        # self.dblargs = np.asarray(dblargs, dtype=np.double)
        # self.boolargs = np.asarray(boolargs, dtype=np.int32)
        # self.longarray = np.asarray(longarray, dtype=np.int_)

    def pack(self):
        """Serialize the data."""
        data_size = self.array.data.itemsize
        if self.array.data.itemsize > ARGS_BUFFER_SIZE:
            raise RuntimeError('Message packet size %d is larger than maximum %d' % (
                data_size, ARGS_BUFFER_SIZE))
        return self.array.data

    def unpack(self, buf):
        """unpack buffer into our data structure."""
        self.array = np.frombuffer(buf, dtype=self.dtype)[0]


def logwrap(func):
    """Decorator for socket send and recv calls, so they can make log."""
    def newfunc(*args, **kwargs):
        logging.debug('%s\t%s\t%s' % (func, args, kwargs))
        try:
            result = func(*args, **kwargs)
        except Exception as ex:
            logging.debug('EXCEPTION: %s' % ex)
            raise
        return result
    return newfunc


class GatanSocket:
    def __init__(self, host, port):
        self.host = host
        self.port = os.environ.get('SERIALEMCCD_PORT', port)
        self.debug = os.environ.get('SERIALEMCCD_DEBUG', 0)
        if self.debug:
            logging.debug('GatanServerIP =', self.host)
            logging.debug('SERIALEMCCD_PORT = GatanServerPort =', self.port)
            logging.debug('SERIALEMCCD_DEBUG =', self.debug)

        self.save_frames = False
        self.num_grab_sum = 0
        self.connect()

        self.script_functions = [
            ('AFGetSlitState', 'GetEnergyFilter'),
            ('AFSetSlitState', 'SetEnergyFilter'),
            ('AFGetSlitWidth', 'GetEnergyFilterWidth'),
            ('AFSetSlitWidth', 'SetEnergyFilterWidth'),
            ('AFDoAlignZeroLoss', 'AlignEnergyFilterZeroLossPeak'),
            ('IFCGetSlitState', 'GetEnergyFilter'),
            ('IFCSetSlitState', 'SetEnergyFilter'),
            ('IFCGetSlitWidth', 'GetEnergyFilterWidth'),
            ('IFCSetSlitWidth', 'SetEnergyFilterWidth'),
            ('IFCDoAlignZeroLoss', 'AlignEnergyFilterZeroLossPeak'),
            ('IFGetSlitIn', 'GetEnergyFilter'),
            ('IFSetSlitIn', 'SetEnergyFilter'),
            ('IFGetEnergyLoss', 'GetEnergyFilterOffset'),
            ('IFSetEnergyLoss', 'SetEnergyFilterOffset'),
            ('IFGetSlitWidth', 'GetEnergyFilterWidth'),
            ('IFSetSlitWidth', 'SetEnergyFilterWidth'),
            ('GT_CenterZLP', 'AlignEnergyFilterZeroLossPeak'),
        ]
        self.filter_functions = {}
        for name, method_name in self.script_functions:
            hasScriptFunction = self.hasScriptFunction(name)
            if self.hasScriptFunction(name):
                self.filter_functions[method_name] = name
            if self.debug:
                logging.debug(name, method_name, hasScriptFunction)
        if ('SetEnergyFilter' in self.filter_functions.keys() and
                self.filter_functions['SetEnergyFilter'] == 'IFSetSlitIn'):
            self.wait_for_filter = 'IFWaitForFilter();'
        else:
            self.wait_for_filter = ''

    def hasScriptFunction(self, name):
        script = 'if ( DoesFunctionExist("%s") ) {{ Exit(1.0); }} else {{ Exit(-1.0); }}' % name
        result = self.ExecuteGetDoubleScript(script)
        return result > 0.0

    def connect(self):
        # recommended by Gatan to use localhost IP to avoid using tcp
        self.sock = socket.create_connection(('127.0.0.1', self.port))

    def disconnect(self):
        self.sock.shutdown(socket.SHUT_RDWR)
        self.sock.close()

    def reconnect(self):
        self.disconnect()
        self.connect()

    @logwrap
    def send_data(self, data):
        return self.sock.sendall(data)

    @logwrap
    def recv_data(self, n):
        return self.sock.recv(n)

    def ExchangeMessages(self, message_send, message_recv=None):
        self.send_data(message_send.pack())

        if message_recv is None:
            return
        recv_buffer = message_recv.pack()
        recv_len = recv_buffer.itemsize

        total_recv = 0
        parts = []
        while total_recv < recv_len:
            remain = recv_len - total_recv
            new_recv = self.recv_data(remain)
            parts.append(new_recv)
            total_recv += len(new_recv)
        buf = b''.join(parts)
        message_recv.unpack(buf)
        # log the error code from received message
        sendargs = message_send.array['longargs']
        recvargs = message_recv.array['longargs']
        logging.debug('Func: %s, Code: %s' % (sendargs[0], recvargs[0]))

    def GetFunction(self, funcName, rlongargs=[], rboolargs=[], rdblargs=[]):
        """ Common function that only receives data. """
        funcCode = enum_gs[funcName]
        message_send = Message(longargs=(funcCode,))
        message_recv = Message(rlongargs, rboolargs, rdblargs)
        self.ExchangeMessages(message_send, message_recv)
        return message_recv

    def SetFunction(self, funcName, slongargs=[], sboolargs=[], sdblargs=[]):
        """ Common function that only sends data. """
        funcCode = enum_gs[funcName]
        message_send = Message(longargs=(funcCode, slongargs, sboolargs, sdblargs))
        message_recv = Message(longargs=(0,))
        self.ExchangeMessages(message_send, message_recv)

    def ExecuteSendCameraObjectionFunction(self, function_name, camera_id=0):
        # first longargs is error code. Error if > 0
        return self.ExecuteGetLongCameraObjectFunction(function_name, camera_id)

    def ExecuteGetLongCameraObjectFunction(self, function_name, camera_id=0):
        """Execute DM script function that requires camera object as input and
        output one long integer."""
        recv_longargs_init = (0,)
        result = self.ExecuteCameraObjectFunction(function_name, camera_id,
                                                  recv_longargs_init=recv_longargs_init)
        if result is False:
            return 1
        return result.array['longargs'][0]

    def ExecuteGetDoubleCameraObjectFunction(self, function_name, camera_id=0):
        """Execute DM script function that requires camera object as input and
        output double floating point number."""
        recv_dblargs_init = (0,)
        result = self.ExecuteCameraObjectFunction(function_name, camera_id,
                                                  recv_dblargs_init=recv_dblargs_init)
        if result is False:
            return -999.0
        return result.array['dblargs'][0]

    def ExecuteCameraObjectFunction(self, function_name, camera_id=0, recv_longargs_init=(0,),
                                    recv_dblargs_init=(0.0,), recv_longarray_init=[]):
        """Execute DM script function that requires camera object as input."""
        if not self.hasScriptFunction(function_name):
            # unsuccessful
            return False
        fullcommand = ('Object manager = CM_GetCameraManager();\n'
                       'Object cameraList = CM_GetCameras(manager);\n'
                       'Object camera = ObjectAt(cameraList,%d);\n'
                       '%s(camera);\n' % (camera_id, function_name))
        result = self.ExecuteScript(fullcommand, camera_id, recv_longargs_init,
                                    recv_dblargs_init, recv_longarray_init)
        return result

    def ExecuteSendScript(self, command_line, select_camera=0):
        recv_longargs_init = (0,)
        result = self.ExecuteScript(command_line, select_camera, recv_longargs_init)
        # first longargs is error code. Error if > 0
        return result.array['longargs'][0]

    def ExecuteGetLongScript(self, command_line, select_camera=0):
        """Execute DM script and return the result as integer."""
        # SerialEMCCD DM TemplatePlugIn::ExecuteScript retval is a double
        return int(self.ExecuteGetDoubleScript(command_line, select_camera))

    def ExecuteGetDoubleScript(self, command_line, select_camera=0):
        """Execute DM script that gets one double float number."""
        recv_dblargs_init = (0.0,)
        result = self.ExecuteScript(command_line, select_camera, recv_dblargs_init=recv_dblargs_init)
        return result.array['dblargs'][0]

    def ExecuteScript(self, command_line, select_camera=0, recv_longargs_init=(0,),
                      recv_dblargs_init=(0.0,), recv_longarray_init=[]):
        funcCode = enum_gs['GS_ExecuteScript']
        cmd_str = command_line + '\0'
        extra = len(cmd_str) % 4
        if extra:
            npad = 4 - extra
            cmd_str = cmd_str + (npad) * '\0'
        # send the command string as 1D longarray
        longarray = np.frombuffer(cmd_str.encode(), dtype=np.int_)
        # logging.debug(longaray)
        message_send = Message(longargs=(funcCode,), boolargs=(select_camera,), longarray=longarray)
        message_recv = Message(longargs=recv_longargs_init, dblargs=recv_dblargs_init,
                               longarray=recv_longarray_init)
        self.ExchangeMessages(message_send, message_recv)
        return message_recv

    def RunScript(self, fn: str, background: bool = False):
        """Run a DM script.

        fn: str
                Path to the script to run
        background: bool
                Prepend `// $BACKGROUND$` to run the script in the background
                and make it non-blocking.
        """

        bkg = r'// $BACKGROUND$\n\n'

        with open(fn, 'r') as f:
            cmd_str = ''.join(f.readlines())

        if background:
            cmd_str = bkg + cmd_str

        return self.ExecuteScript(cmd_str)


class SocketFuncs(GatanSocket):
    def __init__(self, host='127.0.0.1', port=48890):
        super().__init__(host, port)

# ---------- DM functions -----------------------------------------------------

    def GetDMVersion(self):
        message_recv = self.GetFunction('GS_GetDMVersion',
                                        rlongargs=(0, 0))
        result = message_recv.array['longargs'][1]
        return result

    def GetDMVersionAndBuild(self):
        message_recv = self.GetFunction('GS_GetDMVersionAndBuild',
                                        rlongargs=(0, 0, 0))
        result = message_recv.array['longargs']
        return result[0], result[1]

    def GetDMCapabilities(self):
        message_recv = self.GetFunction('GS_GetDMCapabilities',
                                        rlongargs=(0,),
                                        rboolargs=(0, 0, 0))
        result = message_recv.array['boolargs']
        # canSelectShutter, canSetSettling, openShutterWorks
        return list(map(bool, result))

    def GetPluginVersion(self):
        message_recv = self.GetFunction('GS_GetPluginVersion',
                                        rlongargs=(0, 0))
        result = message_recv.array['longargs'][1]
        return result

    def GetLastError(self):
        message_recv = self.GetFunction('GS_GetLastError',
                                        rlongargs=(0, 0))
        result = message_recv.array['longargs'][1]
        return result

    def SetDebugMode(self, mode):
        self.SetFunction('GS_SetDebugMode', slongargs=(mode,))

# ---------- Camera functions -------------------------------------------------

    def GetNumberOfCameras(self):
        message_recv = self.GetFunction('GS_GetNumberOfCameras',
                                        rlongargs=(0, 0))
        result = message_recv.array['longargs'][1]
        return result

    def SetCurrentCamera(self, camera):
        self.SetFunction('GS_SetCurrentCamera', slongargs=(camera,))

    def SelectCamera(self, camera):
        self.SetFunction('GS_SelectCamera',
                         slongargs=(camera,))

    def IsCameraInserted(self, camera):
        funcCode = enum_gs['GS_IsCameraInserted']
        message_send = Message(longargs=(funcCode, camera))
        message_recv = Message(longargs=(0,), boolargs=(0,))
        self.ExchangeMessages(message_send, message_recv)
        result = bool(message_recv.array['boolargs'][0])
        return result

    def InsertCamera(self, camera, state):
        self.SetFunction('GS_InsertCamera',
                         slongargs=(camera,),
                         sboolargs=(state,))

    def SetReadMode(self, mode=-1, scaling=1.0):
        """
        Set the read mode and the scaling factor
        For K2, pass 0, 1, or 2 for linear, counting, super-res
        For K3, pass 3 or 4 for linear or super-res: this is THE signal that it is a K3
        For camera not needing read mode, pass -1
        For OneView, pass -3 for regular imaging or -2 for diffraction
        For K3, the offset to be subtracted for linear mode must be supplied with the scaling.
        The offset is supposed to be 8192 per frame
        The offset per ms is thus nominally (8192 per frame) / (1.502 frames per ms)
        pass scaling = trueScaling + 10 * nearestInt(offsetPerMs)
        """
        self.SetFunction('GS_SetReadMode',
                         slongargs=(mode,),
                         sdblargs=(scaling,))

    def SetShutterNormallyClosed(self, camera, shutter):
        self.SetFunction('GS_SetShutterNormallyClosed',
                         slongargs=(camera, shutter,))

    def GetLastDoseRate(self):
        message_recv = self.GetFunction('GS_GetLastDoseRate',
                                        rlongargs=(0,),
                                        rdblargs=(0,))
        result = float(message_recv.array['dblargs'])
        return result

    @logwrap
    def SetK2Parameters(self, readMode, scaling, hardwareProc, doseFrac,
                        frameTime, alignFrames, saveFrames, filt='', useCds=False):
        funcCode = enum_gs['GS_SetK2Parameters']

        # rotation and flip for non-frame saving image. It is the same definition
        # as in SetFileSaving2
        # if set to 0, it takes what GMS has.self.save_frames = saveFrames
        rotationFlip = 0

        # flags
        flags = 0
        flags += int(useCds) * 2 ** 6
        reducedSizes = 0
        fullSizes = 0

        # filter name
        filt_str = filt + '\0'
        extra = len(filt_str) % 4
        if extra:
            npad = 4 - extra
            filt_str = filt_str + npad * '\0'
        longarray = np.frombuffer(filt_str.encode(), dtype=np.int_)

        longs = [
            funcCode,
            readMode,
            hardwareProc,
            rotationFlip,
            flags
        ]
        bools = [
            doseFrac,
            alignFrames,
            saveFrames,
        ]
        doubles = [
            scaling,
            frameTime,
            reducedSizes,
            fullSizes,
            0.0,  # dummy3
            0.0,  # dummy4
        ]

        message_send = Message(longargs=longs, boolargs=bools,
                               dblargs=doubles, longarray=longarray)
        message_recv = Message(longargs=(0,))  # just return code
        self.ExchangeMessages(message_send, message_recv)

    def setNumGrabSum(self, earlyReturnFrameCount, earlyReturnRamGrabs):
        # pack RamGrabs and earlyReturnFrameCount in one double
        self.num_grab_sum = (2**16) * earlyReturnRamGrabs + earlyReturnFrameCount

    def getNumGrabSum(self):
        return self.num_grab_sum

    @logwrap
    def SetupFileSaving(self, rotationFlip, dirname, rootname, filePerImage,
                        doEarlyReturn, earlyReturnFrameCount=0, earlyReturnRamGrabs=0,
                        lzwtiff=False):
        pixelSize = 1.0
        self.setNumGrabSum(earlyReturnFrameCount, earlyReturnRamGrabs)
        if self.save_frames and (doEarlyReturn or lzwtiff):
            # early return flag
            flag = 128 * int(doEarlyReturn) + 8 * int(lzwtiff)
            numGrabSum = self.getNumGrabSum()
            # set values to pass
            longs = [enum_gs['GS_SetupFileSaving2'], rotationFlip, flag]
            dbls = [pixelSize, numGrabSum, 0., 0., 0.]
        else:
            longs = [enum_gs['GS_SetupFileSaving'], rotationFlip]
            dbls = [pixelSize]
        bools = [filePerImage]
        names_str = dirname + '\0' + rootname + '\0'
        extra = len(names_str) % 4
        if extra:
            npad = 4 - extra
            names_str = names_str + npad * '\0'
        longarray = np.frombuffer(names_str.encode(), dtype=np.int_)
        message_send = Message(longargs=longs, boolargs=bools,
                               dblargs=dbls, longarray=longarray)
        message_recv = Message(longargs=(0, 0))
        self.ExchangeMessages(message_send, message_recv)

    def StopDSAcquisition(self):
        message_recv = self.GetFunction('GS_StopDSAcquisition',
                                        rlongargs=(0, ))

    def StopContinuousCamera(self):
        message_recv = self.GetFunction('GS_StopContinuousCamera',
                                        rlongargs=(0, ))

    def GetFileSaveResult(self):
        message_recv = self.GetFunction('GS_GetFileSaveResult',
                                        rlongargs=(0, 0, 0))
        result = message_recv.array['longargs']
        # result = numSaved, error
        return result[1]

    def SetNoDMSettling(self, value):
        self.SetFunction('GS_SetNoDMSettling',
                         slongargs=(value,))

    def WaitUntilReady(self, value):
        self.SetFunction('GS_WaitUntilReady',
                         slongargs=(value,))

    @logwrap
    def GetImage(self, processing, height, width, binning, top,
                 left, bottom, right, exposure, corrections,
                 shutter=0, shutterDelay=0.):
        """
        :param processing: dark, unprocessed, dark subtracted or gain normalized
        :param exposure: seconds
        :param shutterDelay: milliseconds
        """

        arrSize = width * height

        # TODO: need to figure out what these should be
        divideBy2 = 0
        settling = 0.0

        if processing == 'dark':
            longargs = [enum_gs['GS_GetDarkReference']]
        else:
            longargs = [enum_gs['GS_GetAcquiredImage']]
        longargs.extend([
            arrSize,  # pixels in the image
            width, height
        ])
        if processing == 'unprocessed':
            longargs.append(0)
        elif processing == 'dark subtracted':
            longargs.append(1)
        elif processing == 'gain normalized':
            longargs.append(2)

        longargs.extend([binning, top, left, bottom, right, shutter])
        if processing != 'dark':
            longargs.append(shutterDelay)
        longargs.extend([divideBy2, corrections])
        dblargs = [exposure, settling]

        message_send = Message(longargs=longargs, dblargs=dblargs)
        message_recv = Message(longargs=(0, 0, 0, 0, 0))

        # attempt to solve UCLA problem by reconnecting
        # if self.save_frames:
        # self.reconnect()

        self.ExchangeMessages(message_send, message_recv)

        longargs = message_recv.array['longargs']
        if longargs[0] < 0:
            return 1
        arrSize = longargs[1]
        width = longargs[2]
        height = longargs[3]
        numChunks = longargs[4]
        bytesPerPixel = 2
        numBytes = arrSize * bytesPerPixel
        chunkSize = (numBytes + numChunks - 1) / numChunks
        imArray = np.zeros((height, width), np.ushort)
        received = 0
        remain = numBytes
        for chunk in range(numChunks):
            # send chunk handshake for all but the first chunk
            if chunk:
                message_send = Message(longargs=(enum_gs['GS_ChunkHandshake'],))
                self.ExchangeMessages(message_send)
            thisChunkSize = min(remain, chunkSize)
            chunkReceived = 0
            chunkRemain = thisChunkSize
            while chunkRemain:
                new_recv = self.recv_data(chunkRemain)
                len_recv = len(new_recv)
                imArray.data[received: received + len_recv] = new_recv
                chunkReceived += len_recv
                chunkRemain -= len_recv
                remain -= len_recv
                received += len_recv
        return imArray

    def UpdateK2HardwareDarkReference(self, camera):
        function_name = 'K2_updateHardwareDarkReference'
        return self.ExecuteSendCameraObjectionFunction(function_name, camera)

    def FreeK2GainReference(self, value):
        self.SetFunction('GS_FreeK2GainReference', slongargs=(value,))

    def PrepareDarkReference(self, camera):
        function_name = 'CM_PrepareDarkReference'
        return self.ExecuteSendCameraObjectionFunction(function_name, camera)

# ---------- Energy filter functions ------------------------------------------

    def GetEnergyFilter(self):
        if 'GetEnergyFilter' not in self.filter_functions.keys():
            return -1.0
        func = self.filter_functions['GetEnergyFilter']
        script = 'if ( %s() ) {{ Exit(1.0); }} else {{ Exit(-1.0); }}' % func
        return self.ExecuteGetDoubleScript(script)

    def SetEnergyFilter(self, value):
        if 'SetEnergyFilter' not in self.filter_functions.keys():
            return -1.0
        if value:
            i = 1
        else:
            i = 0
        func = self.filter_functions['SetEnergyFilter']
        wait = self.wait_for_filter
        script = '%s(%d); %s' % (func, i, wait)
        return self.ExecuteSendScript(script)

    def GetEnergyFilterWidth(self):
        if 'GetEnergyFilterWidth' not in self.filter_functions.keys():
            return -1.0
        func = self.filter_functions['GetEnergyFilterWidth']
        script = 'Exit(%s())' % func
        return self.ExecuteGetDoubleScript(script)

    def SetEnergyFilterWidth(self, value):
        if 'SetEnergyFilterWidth' not in self.filter_functions.keys():
            return -1.0
        func = self.filter_functions['SetEnergyFilterWidth']
        script = 'if ( %s(%f) ) {{ Exit(1.0); }} else {{ Exit(-1.0); }}' % (func, value)
        return self.ExecuteSendScript(script)

    def GetEnergyFilterOffset(self):
        if 'GetEnergyFilterOffset' not in self.filter_functions.keys():
            return 0.0
        func = self.filter_functions['GetEnergyFilterOffset']
        script = 'Exit(%s())' % func
        return self.ExecuteGetDoubleScript(script)

    def SetEnergyFilterOffset(self, value):
        if 'SetEnergyFilterOffset' not in self.filter_functions.keys():
            return -1.0
        func = self.filter_functions['SetEnergyFilterOffset']
        script = 'if ( %s(%f) ) {{ Exit(1.0); }} else {{ Exit(-1.0); }}' % (func, value)
        return self.ExecuteSendScript(script)

    def AlignEnergyFilterZeroLossPeak(self):
        func = self.filter_functions['AlignEnergyFilterZeroLossPeak']
        wait = self.wait_for_filter
        script = ' if ( %s() ) {{ %s Exit(1.0); }} else {{ Exit(-1.0); }}' % (func, wait)
        return self.ExecuteGetDoubleScript(script)
