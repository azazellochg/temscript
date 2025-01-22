import numpy as np
import json
import sys
import base64
import zlib
import gzip
import io
import functools
import logging


MIME_TYPE_PICKLE = "application/python-pickle"
MIME_TYPE_JSON = "application/json"


class ExtendedJsonEncoder(json.JSONEncoder):
    """JSONEncoder which handles iterables and numpy types"""
    def default(self, obj):
        if isinstance(obj, np.generic):
            return obj.item()
        try:
            iterable = iter(obj)
        except TypeError:
            pass
        else:
            return list(iterable)
        return super(ExtendedJsonEncoder, self).default(obj)


ARRAY_TYPES = {
    "INT8": np.int8,
    "INT16": np.int16,
    "INT32": np.int32,
    "INT64": np.int64,
    "UINT8": np.uint8,
    "UINT16": np.uint16,
    "UINT32": np.uint32,
    "UINT64": np.uint64,
    "FLOAT32": np.float32,
    "FLOAT64": np.float64
}

ARRAY_ENDIANNESS = {"LITTLE", "BIG"}


def unpack_array(obj):
    """
    Unpack a packed array.

    :param obj: Dict with packed array
    """
    sys_endianness = sys.byteorder.upper()
    shape = int(obj["height"]), int(obj["width"])
    dtype = ARRAY_TYPES[obj["type"]]
    endianess = obj["endianness"]
    if endianess not in ARRAY_ENDIANNESS:
        raise ValueError("Unsupported endianness for encoded array: %s" % str(endianess))
    encoding = obj["encoding"]
    if encoding == "BASE64":
        data = base64.b64decode(obj["data"])
    else:
        raise ValueError("Unsupported encoding of array in JSON stream: %s" % str(encoding))
    data = np.frombuffer(data, dtype=dtype).reshape(*shape)
    if obj["endianness"] != sys_endianness:
        data = data.byteswap()
    return data


def pack_array(array):
    """
    Pack array for JSON serialization.

    :param array: Numpy array to pack
    """
    array = np.asanyarray(array)

    type_name = array.dtype.name.upper()
    if type_name not in ARRAY_TYPES:
        raise TypeError("Array data type %s can not be packed" % type_name)

    if array.dtype.byteorder == '<':
        endianness = "LITTLE"
    elif array.dtype.byteorder == '>':
        endianness = "BIG"
    else:
        endianness = sys.byteorder.upper()

    return {
        'width': array.shape[1],
        'height': array.shape[0],
        'type': type_name,
        'endianness': endianness,
        'encoding': "BASE64",
        'data': base64.b64encode(array).decode("ascii")
    }


def gzip_encode(content):
    """GZIP encode bytes object"""
    out = io.BytesIO()
    f = gzip.GzipFile(fileobj=out, mode='w', compresslevel=5)
    f.write(content)
    f.close()
    return out.getvalue()


def gzip_decode(content):
    """Decode GZIP encoded bytes object"""
    return zlib.decompress(content, 16 + zlib.MAX_WBITS)    # No keyword arguments until Python 3.6


def rgetattr(obj, attr, *args, **kwargs):
    """ Recursive getattr or callable on a COM object"""
    try:
        if kwargs.get("typeget", True):
            logging.info("Getattr: %s, args=%s, kwargs=%s" %
                         (attr, args, kwargs))
        result = functools.reduce(getattr, attr.split('.'), obj)
        return result(*args, **kwargs) if callable(result) else result
    except AttributeError:
        logging.error("Attribute error %s" % attr)
        raise AttributeError("AttributeError: %s" % attr)

def rsetattr(obj, attr, val):
    """ https://stackoverflow.com/a/31174427 """
    logging.info("Setattr: %s=%s" % (attr, val))
    pre, _, post = attr.rpartition('.')
    return setattr(rgetattr(obj, pre, typeget=False) if pre else obj, post, val)
