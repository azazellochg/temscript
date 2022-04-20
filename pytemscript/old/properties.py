
class BaseProperty:
    __slots__ = '_com_obj', '_name', '_readonly'

    def __init__(self, com_obj, name='', readonly=False):
        self._com_obj = com_obj
        self._name = name
        self._readonly = readonly

#    def __set_name__(self, owner, name):
#        self._name = " '%s'" % name


class EnumProperty(BaseProperty):
    __slots__ = '_enum_type'

    def __init__(self, com_obj, enum_type, name='', readonly=False):
        super().__init__(com_obj=com_obj, name=name, readonly=readonly)
        self._enum_type = enum_type

    def __get__(self, obj, objtype=None):
        value = getattr(self._com_obj, self._name)
        return self._enum_type(value).name

    def __set__(self, obj, value):
        if self._readonly:
            raise AttributeError("Attribute %s is not writable" % self._name)
        setattr(self._com_obj, self._name, int(value))

    #def __set_name__(self, owner, name):
    #    self._name = " '%s'" % name


class VectorProperty(BaseProperty):

    def __get__(self, obj, objtype=None):
        result = getattr(self._com_obj, self._name)
        return result.X, result.Y

    def __set__(self, obj, value):
        if self._readonly:
            raise AttributeError("Attribute %s is not writable" % self._name)
        value = [float(c) for c in value]
        if len(value) != 2:
            raise ValueError("Expected two items for attribute %s." % self._name)

        vector = getattr(self._com_obj, self._name)
        vector.X = value[0]
        vector.Y = value[1]

        setattr(self._com_obj, self._name, vector)

'''
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
            NewCollectionProperty.GET_ITEM_METHOD(collection.get(), n, item.byref())
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
            raise AttributeError("Attribute %s is not readable" % self._name)
        result = FegFocusIndex()
        prototype = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(self._get_index, "get_property")
        prototype(obj.get(), result.byref())
        return result.Coarse, result.Fine

'''
