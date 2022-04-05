import ctypes
from temscript.com.com_base import *


class BaseProperty:
    __slots__ = '_get_index', '_put_index', '_name'

    def __init__(self, get_index=None, put_index=None):
        self._get_index = get_index
        self._put_index = put_index
        self._name = ''

    def __set_name__(self, owner, name):
        self._name = " '%s'" % name


class LongProperty(BaseProperty):
    __slots__ = '_get_index', '_put_index', '_name'

    def __get__(self, obj, objtype=None):
        if self._get_index is None:
            raise AttributeError("Attribute %sis not readable" % self._name)
        result = ctypes.c_long(-1)
        prototype = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(self._get_index, "get_property")
        prototype(obj.get(), ctypes.byref(result))
        return result.value

    def __set__(self, obj, value):
        if self._put_index is None:
            raise AttributeError("Attribute %sis not writable" % self._name)
        value = int(value)
        prototype = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_long)(self._put_index, "put_property")
        prototype(obj.get(), value)


class VariantBoolProperty(BaseProperty):
    def __get__(self, obj, objtype=None):
        if self._get_index is None:
            raise AttributeError("Attribute %sis not readable" % self._name)
        result = ctypes.c_short(-1)
        prototype = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(self._get_index, "get_property")
        prototype(obj.get(), ctypes.byref(result))
        return bool(result.value)

    def __set__(self, obj, value):
        if self._put_index is None:
            raise AttributeError("Attribute %sis not writable" % self._name)
        bool_value = 0xffff if value else 0x0000
        prototype = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_short)(self._put_index, "put_property")
        prototype(obj.get(), bool_value)


class DoubleProperty(BaseProperty):
    def __get__(self, obj, objtype=None):
        if self._get_index is None:
            raise AttributeError("Attribute %sis not readable" % self._name)
        result = ctypes.c_double(-1)
        prototype = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(self._get_index, "get_property")
        prototype(obj.get(), ctypes.byref(result))
        return result.value

    def __set__(self, obj, value):
        if self._put_index is None:
            raise AttributeError("Attribute %sis not writable" % self._name)
        value = float(value)
        prototype = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_double)(self._put_index, "put_property")
        prototype(obj.get(), value)


class StringProperty(BaseProperty):
    def __get__(self, obj, objtype=None):
        if self._get_index is None:
            raise AttributeError("Attribute %sis not readable" % self._name)
        result = BStr()
        prototype = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(self._get_index, "get_property")
        prototype(obj.get(), result.byref())
        return result.value

    def __set__(self, obj, value):
        if self._put_index is None:
            raise AttributeError("Attribute %sis not writable" % self._name)
        value = BStr(str(value))
        prototype = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(self._put_index, "put_property")
        prototype(obj.get(), BStr(value).get())


class EnumProperty(BaseProperty):
    __slots__ = '_enum_type'

    def __init__(self, enum_type, get_index=None, put_index=None):
        super(EnumProperty, self).__init__(get_index=get_index, put_index=put_index)
        self._enum_type = enum_type

    def __get__(self, obj, objtype=None):
        if self._get_index is None:
            raise AttributeError("Attribute %sis not readable" % self._name)
        result = ctypes.c_int(-1)
        prototype = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(self._get_index, "get_property")
        prototype(obj.get(), ctypes.byref(result))
        return self._enum_type(result.value)

    def __set__(self, obj, value):
        if self._put_index is None:
            raise AttributeError("Attribute %sis not writable" % self._name)
        value = int(value)
        prototype = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_long)(self._put_index, "put_property")
        prototype(obj.get(), value)

    def __set_name__(self, owner, name):
        self._name = " '%s'" % name


class Vector(IUnknown):
    IID = UUID("9851bc47-1b8c-11d3-ae0a-00a024cba50c")

    X = DoubleProperty(get_index=7, put_index=8)
    Y = DoubleProperty(get_index=9, put_index=10)


class VectorProperty(BaseProperty):
    __slots__ = '_get_prototype'

    def __init__(self, get_index, put_index=None):
        super(VectorProperty, self).__init__(get_index=get_index, put_index=put_index)
        self._get_prototype = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(get_index, "get_property")

    def __get__(self, obj, objtype=None):
        result = Vector()
        self._get_prototype(obj.get(), result.byref())
        return result.X, result.Y

    def __set__(self, obj, value):
        if self._put_index is None:
            raise AttributeError("Attribute%s is not writable" % self._name)

        value = [float(c) for c in value]
        if len(value) != 2:
            raise ValueError("Expected two items for attribute%s." % self._name)

        result = Vector()
        self._get_prototype(obj.get(), result.byref())
        result.X = value[0]
        result.Y = value[1]
        prototype = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(self._put_index, "put_property")
        prototype(obj.get(), result.get())


class ObjectProperty(BaseProperty):
    __slots__ = '_interface'

    def __init__(self, interface, get_index, put_index=None):
        super(ObjectProperty, self).__init__(get_index=get_index, put_index=put_index)
        self._interface = interface

    def __get__(self, obj, objtype=None):
        result = self._interface()
        prototype = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(self._get_index, "get_property")
        prototype(obj.get(), result.byref())
        return result

    def __set__(self, obj, value):
        if self._put_index is None:
            raise AttributeError("Attribute%s is not writable" % self._name)
        if not isinstance(value, self._interface):
            raise TypeError("Expected attribute%s to be set to an instance of type %s" % (self._name, self._interface.__name__))
        prototype = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(self._put_index, "put_property")
        prototype(obj.get(), value.get())


class CollectionProperty(BaseProperty):
    __slots__ = '_interface'

    GET_COUNT_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(7, "get_Count")
    GET_ITEM_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, VARIANT, ctypes.c_void_p)(8, "get_Item")

    def __init__(self, get_index, interface=None):
        super(CollectionProperty, self).__init__(get_index=get_index)
        if interface is None:
            interface = IUnknown
        self._interface = interface

    def __get__(self, obj, objtype=None):
        collection = IUnknown()
        prototype = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(self._get_index, "get_property")
        prototype(obj.get(), collection.byref())

        count = ctypes.c_long(-1)
        CollectionProperty.GET_COUNT_METHOD(collection.get(), ctypes.byref(count))
        result = []

        for n in range(count.value):
            index = Variant(n, vartype=VariantType.I4)
            item = self._interface()
            CollectionProperty.GET_ITEM_METHOD(collection.get(), index.get(), item.byref())
            result.append(item)

        return result


class NewCollectionProperty(BaseProperty):
    """ Advanced scripting uses c_long for item index. """
    __slots__ = '_interface'

    GET_COUNT_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(9, "get_Count")
    GET_ITEM_METHOD = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_long, ctypes.c_void_p)(8, "get_Item")

    def __init__(self, get_index, interface=None):
        super(NewCollectionProperty, self).__init__(get_index=get_index)
        if interface is None:
            interface = IUnknown
        self._interface = interface

    def __get__(self, obj, objtype=None):
        collection = IUnknown()
        prototype = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(self._get_index, "get_property")
        prototype(obj.get(), collection.byref())

        count = ctypes.c_long(-1)
        NewCollectionProperty.GET_COUNT_METHOD(collection.get(), ctypes.byref(count))
        result = []

        for n in range(count.value):
            item = self._interface()
            NewCollectionProperty.GET_ITEM_METHOD(collection.get(), ctypes.c_long(n), item.byref())
            result.append(item)

        return result


class SafeArrayProperty(BaseProperty):
    def __get__(self, obj, objtype=None):
        result = SafeArray()
        prototype = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(self._get_index, "get_property")
        prototype(obj.get(), result.byref())
        return result


class FegFocusIndexProperty(BaseProperty):
    __slots__ = '_get_index', '_name'

    def __get__(self, obj, objtype=None):
        if self._get_index is None:
            raise AttributeError("Attribute %sis not readable" % self._name)
        result = FegFocusIndex()
        prototype = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(self._get_index, "get_property")
        prototype(obj.get(), ctypes.byref(result))
        return result.Coarse, result.Fine
